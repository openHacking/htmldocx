import { Document, Packer, Paragraph, TextRun } from "docx";

export async function saveDocx(htmlString) {
    const parser = new DOMParser();
    const docHTML = parser.parseFromString(htmlString, 'text/html');


    const paragraphs = []
    const nodes = docHTML.getElementsByTagName('*');


    for (let i = 0; i < nodes.length; i++) {
        const node = nodes[i];
        let paragraph = convertNode(node)
        if (paragraph) {
            paragraphs.push(paragraph)
        }
    }

    const doc = new Document({
        creator: "Alex",
        description: "My GPTdocument",
        title: "GPT Document",
        sections: [{
            properties: {},
            children: paragraphs
        }]
    });


    // saveBuffer(doc)
    saveBlob(doc)
}


function convertNode(node) {
    let element;
    const tagName = node.tagName.toLowerCase();
    switch (tagName) {
        case 'p':
            element = convertP(node);
            break;
        case 'h1':
            element = convertHeading(node, 1);
            break;
        case 'h2':
            element = convertHeading(node, 2);
            break;
        case 'h3':
            element = convertHeading(node, 3);
            break;
        case 'h4':
            element = convertHeading(node, 4);
            break;
        case 'h5':
            element = convertHeading(node, 5);
            break;
        case 'h6':
            element = convertHeading(node, 6);
            break;
        // 处理其他标签
        default:
            // element = convertDefault(node);
            break;
    }
    return element;
}


function convertP(node) {
    const element = new Paragraph({
        children: [
            new TextRun(node.innerText),
        ],
    })
    return element;
}
function convertHeading(node, level) {
    const element = new Paragraph({
        text: node.innerText,
        heading: level,
    })
    return element;
}


async function saveBuffer(doc) {
    const packer = new Packer();
    const buffer = await packer.toBuffer(doc);
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.click()
    a.remove()
}
function saveBlob(doc) {
    Packer.toBlob(doc).then((blob) => {
        // saveAs from FileSaver will download the file
        downloadBlob(blob, {
            fileName: "gpt",
            extName: "docx",
        });
    });
}

/**
 * download blob as file
 * @param {*} blob
 */
export const downloadBlob = (blob, args) => {
    const fileName = `${args.fileName || new Date().valueOf()}.${args.extName || "txt"
        }`;
    const link = document.createElement("a");
    link.href = window.URL.createObjectURL(blob);
    link.download = fileName;
    link.click();
    window.URL.revokeObjectURL(link.href);
};