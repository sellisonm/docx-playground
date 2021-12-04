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


let createTableP = function(dataRows) {

    let t = new docx.Table({
        columnWidths: [3505, 5505],
        borders: {
            top: {
                style: docx.BorderStyle.DASH_DOT_STROKED,
                size: 1,
                color: "ff0000",
            },
            bottom: {
                style: docx.BorderStyle.THICK_THIN_MEDIUM_GAP,
                size: 5,
                color: "889900",
            },
        },
        rows: [new docx.TableRow({
            children: [new docx.TableCell({
                children: [new docx.Paragraph("test")]
            })]
        })]
    })

    dataRows.forEach(dr => {

        let tr = new docx.TableRow({
            children: [new docx.TableCell({
                children: [new docx.Paragraph(dr.val)]
            })]
        })
       // let tc = 
        //let tp = new docx.Paragraph(dr.val)

        //tc.children.push(tp)
        //tr.children.push(tc)
        
    })

    return t
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

    if (Array.isArray(paras)) {
        paras.unshift(p)
    }
    else {
        paras = new Array(new docx.Table({
        columnWidths: [3505, 5505],
        borders: {
            top: {
                style: docx.BorderStyle.DASH_DOT_STROKED,
                size: 1,
                color: "ff0000",
            },
            bottom: {
                style: docx.BorderStyle.THICK_THIN_MEDIUM_GAP,
                size: 5,
                color: "889900",
            },
        },
        rows: [new docx.TableRow({
                children: [new docx.TableCell({
                 children: [new docx.Paragraph("test")]
                })]
            })]
            })
        )
    }

    let section = {
        properties: {
            type: docx.SectionType.CONTINUOUS
        },
        children: [new docx.Table({
            columnWidths: [3505, 5505],
            borders: {
                top: {
                    style: docx.BorderStyle.DASH_DOT_STROKED,
                    size: 1,
                    color: "ff0000",
                },
                bottom: {
                    style: docx.BorderStyle.THICK_THIN_MEDIUM_GAP,
                    size: 5,
                    color: "889900",
                },
            },
            rows: [new docx.TableRow({
                    children: [new docx.TableCell({
                     children: [new docx.Paragraph("test")]
                    })]
                })]
            })
        ]
            
    }    
    return section
}

let sectionsList = []

for(const section in data) {

        let paras = []
        let tablePara = null
        let thisSection 

        if (Object.hasOwnProperty.call(data, section)) {
            const el = data[section];

            for(const q in el.questions)
            {
                if (Object.hasOwnProperty.call(el.questions, q)) {
                    const qel = el.questions[q];

                    // let questionPara = createQuestionP(qel.text)
                    // paras.push(questionPara)

                    if (qel.type === "table") {
                        let fakeDataRow = []
                        fakeDataRow.push({val: "5.3"})
                        fakeDataRow.push({val: "9.2"})
                        tablePara = createTableP(fakeDataRow)                        
                    }
                    // else {                    
                    //     let answerPara = createAnswerP(qel.answer)
                    //     paras.push(answerPara)

                    //     if (qel.type === "dropdown") {
                    //         paras.push(createOptionsP(qel.options))
                    //     }
                    // }
                }
            }

            if (Object.hasOwnProperty.call(data, section)) {
                const el = data[section];

                if (tablePara != null) {
                    thisSection = createTrueS(el.name, tablePara)
                 }
                // else {
                //     thisSection = createTrueS(el.name, paras)
                // }
                // tablePara = null
            }

            if (thisSection)
                sectionsList.push(thisSection)
    }
}


const doc = new docx.Document({
    sections: sectionsList
});

// Used to export the file into a .docx file
docx.Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});

// Done! A file called 'My Document.docx' will be in your file system.