const fs = require('fs')
const path = require('path')
const http = require('http')
const officegen = require('officegen')


const handleFile = (f) => {
    const contentText = fs.readFileSync(path.join(__dirname, './data/' + f), 'utf-8')
    const analysisText = fs.readFileSync(path.join(__dirname, './analysis/' + f), 'utf-8')
    const doc = fs.createWriteStream(path.join(__dirname, './doc/' + f.split('.')[0] + '.docx'))


    try {
        if(!contentText || !analysisText) {
            console.log(`${f}json未读取`)
            throw new Error()
        }

        let contentJson = null
        let analysisJson = null

        try {
            contentJson = JSON.parse(contentText)
            analysisJson = JSON.parse(analysisText)
        } catch (e) {
            console.log(`${f}json反序列化失败`)
            throw new Error(e)
        }


        const docConfig = {
            'type': 'docx',
            'subject': f,
            'keywords': '',
            'description': ''
        }
        const docx = officegen (docConfig)

        //题目数
        let index = 1
        const ENUM = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

        const compose = (...funcs) => {
            if(funcs.length === 0) {
                return args => args
            }
            if(funcs.length === 1) {
                return funcs[0]
            }
            return funcs.reduce((a, b) => (...args) => a(b(...args)))
        }

        const downloadImg = (file, uri) => {
            return new Promise((resolve, reject) => {
                http.get(uri, (res) => {
                    let imgData = ''

                    res.setEncoding("binary")

                    res.on('data', (data) => {
                        imgData += data
                    }).on('end', () => {
                        fs.writeFileSync(file, imgData, "binary")
                    })
                })
            })
        }

        const getUrlParams = (url) => {
            const paramsObj = {}
            const arrUrl = url.split("?")
            const paramsStr = arrUrl[1]
            if(arrUrl[1]) {
                const arrParamsStr = paramsStr.split("&")
                arrParamsStr.map((v) => {
                    const arr = v.split("=")
                    paramsObj[arr[0]] = arr[1]
                })
            }
            return paramsObj
        }

        //处理图片
        const todoImg = async (value, pObj) => {
            const uri = `http://fb.fbstatic.cn/api/xingce/images/${value}`

            const outPath = path.join(__dirname, `./tmp/${value.split('?')[0]}`)
            const exists = fs.existsSync(outPath)

            if(!exists) {
                await downloadImg(outPath, uri)
            } else {
                //const { width, height } = getUrlParams(value)
                pObj.addLineBreak()
                pObj.addImage(outPath)
            }

        }

        //递归children函数
        const recursionChildren = (children, pObj) => {
            if('children' in children) {
                children['children'].map(c => {
                    switch (c.name) {
                        case 'img':
                            todoImg(c.value, pObj)
                            break
                        case 'tex': //数学公式
                            pObj.addText('[数学公式]', {
                                align: 'left',
                                font_face: 'Arial',
                                font_size: 14
                            })
                            break
                        case 'u':
                            pObj.addText('____', {
                                align: 'left',
                                font_face: 'Arial',
                                font_size: 14
                            })
                            break
                        default:
                            pObj.addText(c.value, {
                                align: 'left',
                                font_face: 'Arial',
                                font_size: 14
                            })
                            break
                    }

                    if(c.children.length !== 0) {
                        recursionChildren(c, pObj)
                    }
                })
            }
        }

        //处理题目素材
        const handleMaterial = ({json, pObj}) => {
            if(!json.material) {
                return {json, pObj}
            } else {
                const _material = JSON.parse(json.material.content)
                recursionChildren(_material, pObj)
                pObj.addLineBreak()
                pObj.addLineBreak()

                return {json, pObj}
            }
        }

        //处理题目内容
        const handleContent = ({json, pObj}) => {
            if(!json.content) {
                return {json, pObj}
            } else {
                const _content = JSON.parse(json.content)

                pObj.addText(`${index}.`, {
                    align: 'left',
                    font_face: 'Arial',
                    font_size: 14
                })

                recursionChildren(_content, pObj)
                return {json, pObj}
            }
        }

        //处理题目答案
        const handleAnswer = ({json, pObj}) => {
            const choice = parseInt(json.correctAnswer.choice)
            const _pObj = docx.createP()
            _pObj.addText(`正确答案(${ENUM[choice]})`, {
                align: 'left',
                font_face: 'Arial',
                font_size: 14
            })
            return {json, pObj}
        }

        //处理题目选项
        const handleAccessories = ({json, pObj}) => {
            if(!json.accessories) {
                return {json, pObj}
            } else {
                const _accessories = json.accessories
                const accessories = _accessories[0].options

                accessories.map((a, i) => {
                    const pObj = docx.createP()

                    try {
                        const _acs = JSON.parse(a),
                            acs = _acs.children

                        pObj.addText(ENUM[i] + '.', {
                            align: 'left',
                            font_face: 'Arial',
                            font_size: 14
                        })

                        acs.map((a, i) => {
                            recursionChildren(a, pObj)
                        })

                    } catch(e) {
                        pObj.addText(ENUM[i] + '.' + a, {
                            align: 'left',
                            font_face: 'Arial',
                            font_size: 14
                        })
                    }
                })

                return {json, pObj}
            }
        }

        //处理每题解析
        const handleAnalysis = ({json, pObj}) => {
            const solution = JSON.parse(analysisJson.filter(a => a.id === json.id)[0]['solution'])

            const _pObj = docx.createP()
            _pObj.addText('解析:', {
                align: 'left',
                font_face: 'Arial',
                font_size: 14
            })
            recursionChildren(solution, _pObj)

            return {json, pObj}
        }

        //每题间隔
        const handleSpace = () => {
            const pObj = docx.createP()
            pObj.addLineBreak()
        }

        contentJson.map(c => {
            compose(
                handleSpace,
                handleAnalysis,
                handleAnswer,
                handleAccessories,
                handleContent,
                handleMaterial,
            )({json: c, pObj: docx.createP()})

            index++
        })

        const result = docx.generate(doc, {
            'finalize': function ( written ) {
                console.log ( 'Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n' );
            },
            'error': function ( err ) {
                console.log ( err );
            }
        })
    } catch (e) {
        console.log(`Error in ${f}`)
    }
}

//handleFile('2018年421联考《行测》真题（天津卷）（网友回忆版）.json')

const files =fs.readdirSync(path.join(__dirname, './data/'))
files.map(f => handleFile(f))
