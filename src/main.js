const docx = require("docx")
const fs = require("fs")
const data = require("./data/data.json")

let createQuestionP = function(t) {
    let p = new docx.Paragraph({
        heading: docx.HeadingLevel.HEADING_3,
        spacing: {
            before: 300
        },
        children: [
            new docx.TextRun(t)
        ]
    })

    return p
}
let createAnswerP = function(t) {
    let p = new docx.Paragraph({ 
        keepLines: true,       
        children: [
            new docx.TextRun(t)
        ]
    })

    return p
}
let createOptionsP = function(t) {
    t = "Possible Options: " + t
    let p = new docx.Paragraph({ 
        keepLines: true, 

        children: [
            new docx.TextRun({
                text: t,
                italics: true,
                size: 16,
                color: "585858"
            })
        ]
    })

    return p
}
let createS = function(newLine, p) {
    let section = {
        properties: {
            type: docx.SectionType.CONTINUOUS
        },
        children: [ 
            p
        ]
    }
    
    return section
}

let createTrueS = function(name, paras) {
    let p = new docx.Paragraph({
        heading: docx.HeadingLevel.HEADING_1,
        spacing: {
            before: 500
        },
        children: [
            new docx.TextRun(name)
        ]
    })

    paras.unshift(p)

    let section = {
        properties: {
            type: docx.SectionType.CONTINUOUS
        },
        children: paras
    }    
    return section
}

//let questionChildren = []
let sectionsList = []

for(const section in data) {

        let paras = []
        let thisSection 

        if (Object.hasOwnProperty.call(data, section)) {
            const el = data[section];

            for(const q in el.questions)
            {
                if (Object.hasOwnProperty.call(el.questions, q)) {
                    const qel = el.questions[q];
                    let questionPara = createQuestionP(qel.text)
                    //sectionsList.push(createS(true, questionPara))
                    let answerPara = createAnswerP(qel.answer)
                    //sectionsList.push(createS(false, answerPara))
                    paras.push(questionPara, answerPara)
                    if (qel.type === "dropdown") {
                        paras.push(createOptionsP(qel.options))
                    }
                }
            }


            if (Object.hasOwnProperty.call(data, section)) {
                const el = data[section];
                thisSection = createTrueS(el.name, paras)
            }

            sectionsList.push(thisSection)
        }

        const t = new docx.Table({
            columnWidths: [3505, 5505],
            rows: [
                new docx.TableRow({
                    children: [
                        new docx.TableCell({
                            width: {
                                size: 3505,
                                type: docx.WidthType.DXA,
                            },
                            children: [new docx.Paragraph("Hello")],
                        }),
                        new docx.TableCell({
                            width: {
                                size: 5505,
                                type: docx.WidthType.DXA,
                            },
                            children: [new docx.Paragraph("World")],
                        }),
                    ],
                }),
            ],
        });


        sectionsList.push({
            properties: {
                type: docx.SectionType
            },
            children: [t]
        })
}


const doc = new docx.Document({
    sections: sectionsList
});

// Used to export the file into a .docx file
docx.Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});

// Done! A file called 'My Document.docx' will be in your file system.