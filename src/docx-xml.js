
/**
 * http://officeopenxml.com/WPsampleDoc.php
 */

// 导入jszip库
import JSZip from 'jszip';


// 生成docx文件的函数
export class GenerateDocx {
    constructor(html){
        this.html = html;

        this.hyperlinkXML = ``;
        this.hyperlinkId = 8;
        this.init()
    }

    init(){

    // 解析 HTML 字符串为 DOM 对象
    const parser = new DOMParser();
    const doc = parser.parseFromString(this.html, "text/html");

    // 创建一个新的JSZip实例
    const zip = this.zip = new JSZip();

      let xml = ``;
      
    for (const node of doc.body.childNodes) {
        xml += this.traverse(node);
    }
    // 将构建的xml字符串添加到JSZip实例中
    setContentTypes(zip)
    setRels(zip)
    setCustomXml(zip)
    setDocProps(zip)
    setWord(zip, xml,this.hyperlinkXML)
    }

    traverse(node,setting) {
        // 处理文本节点
        if (node.nodeType === Node.TEXT_NODE) {
            return node.textContent.trim() === '' ? '' : `<w:r><w:t>${node.textContent}</w:t></w:r>`;
        }
    
        // 处理加粗标签，行内代码
        else if (
            node.nodeType === Node.ELEMENT_NODE &&
            (node.nodeName === "STRONG" || node.nodeName === "B" || node.nodeName === "CODE")
        ) {
            let childrenXml = "";
            for (const child of node.childNodes) {
                childrenXml += this.traverse(child);
            }
            return `<w:r><w:rPr><w:b/></w:rPr>${childrenXml}</w:r>`;
        }
        // 处理斜体标签
        else if (
            node.nodeType === Node.ELEMENT_NODE &&
            (node.nodeName === "EM" || node.nodeName === "I")
        ) {
            let childrenXml = "";
            for (const child of node.childNodes) {
                childrenXml += this.traverse(child);
            }
            return `<w:r><w:rPr><w:i/></w:rPr>${childrenXml}</w:r>`;
        }
    
        // 处理段落标签
        else if (node.nodeType === Node.ELEMENT_NODE && node.nodeName === "P") {
            let childrenXml = "";
            for (const child of node.childNodes) {
                childrenXml += this.traverse(child);
            }

            // 引用中用到了P。，但是不需要用<w:p>包裹
            return setting && setting.isRemoveP ? childrenXml :`<w:p>${childrenXml}</w:p>`;
        }
        // 处理链接标签
        else if (node.nodeType === Node.ELEMENT_NODE && node.nodeName === "A") {
            let childrenXml = "";
            for (const child of node.childNodes) {
                childrenXml += this.traverse(child);
            }
            const rId = `rId${this.hyperlinkId++}`
            this.hyperlinkXML += `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://chatopenai.pro/" TargetMode="External"/>`
            return `<w:hyperlink r:id="${rId}" w:history="1">
                        <w:r>
                            <w:rPr>
                                <w:rStyle w:val="Hyperlink"/>
                            </w:rPr>
                            ${childrenXml}
                        </w:r>
                    </w:hyperlink>`;
        }
    
        // 处理 H1-H6 标签
        else if (node.nodeType === Node.ELEMENT_NODE && /^H[1-6]$/.test(node.nodeName)) {
            const headingLevel = node.nodeName.substring(1);
            let childrenXml = "";
            for (const child of node.childNodes) {
                childrenXml += this.traverse(child);
            }
            return `<w:p><w:pPr><w:pStyle w:val="Heading${headingLevel}" /></w:pPr>${childrenXml}</w:p>`;
        }
    
        // 处理无序列表、有序列表标签
        else if (node.nodeType === Node.ELEMENT_NODE && (node.nodeName === "UL" || node.nodeName === "OL")) {
            let childrenXml = "";
            const childNodes = Array.from(node.childNodes).filter(n => n.nodeName === 'LI')

            let numId = node.nodeName === "UL" ? 5 : 6;
            let listLevel = 0;
            // 嵌套的列表，识别到0,1,2...就要递增等级
            if(setting && (setting.listLevel || setting.listLevel === 0)){
                listLevel++
            }
            for (const child of childNodes) {
                childrenXml += `<w:p><w:pPr>
                <w:pStyle w:val="ListParagraph"/>
                <w:numPr>
                    <w:ilvl w:val="${listLevel}"/>
                    <w:numId w:val="${numId}"/>
                </w:numPr>
                <w:ind w:firstLineChars="0"/>
            </w:pPr>${this.traverse(child)}</w:p>`;
            }
            return childrenXml;
        }

        // 处理文本节点
        if (node.nodeType === Node.ELEMENT_NODE && node.nodeName === "LI") {
                return `<w:r><w:t>${node.textContent}</w:t></w:r>`;
        }
    
    
    
        // 处理代码块标签
        else if (node.nodeType === Node.ELEMENT_NODE && node.nodeName === "PRE") {
            const languageXml = `<w:t>${node.firstChild.firstChild.innerText}</w:t><w:br/>`;
            const codeXml = node.firstChild.lastChild.innerText.split('\n').map(str => `<w:t xml:space="preserve">${escapeXml(str)}</w:t>`).join('<w:br/>')
            return `<w:p><w:r>${languageXml}${codeXml}</w:r></w:p>`
        }
        // 处理分割线标签
        else if (node.nodeType === Node.ELEMENT_NODE && node.nodeName === "HR") {
            return `<w:p>
            <w:pPr>
                <w:pBdr>
                    <w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>
                </w:pBdr>
            </w:pPr>
            <w:r>
                <w:lastRenderedPageBreak/>
            </w:r>
        </w:p>`
        }
        // 处理引用标签
        else if (node.nodeType === Node.ELEMENT_NODE && node.nodeName === "BLOCKQUOTE") {
            let childrenXml = "";
            for (const child of node.childNodes) {
                childrenXml += this.traverse(child,{isRemoveP:true});
            }

            return `<w:p>
            <w:pPr>
                <w:pStyle w:val="Quote"/>
                <w:jc w:val="left"/>
            </w:pPr>
            ${childrenXml}
        </w:p>`
        }
        // 处理表格标签
        else if (node.nodeType === Node.ELEMENT_NODE && node.nodeName === "TABLE") {
            return parseTable(node)
        }
    
        // 处理其他标签
        // ...
    
        return "";
    }

    save(name = 'docx-template'){
        // 将生成的docx文件下载到本地
        this.zip.generateAsync({ type: 'blob' }).then(function (content) {
            const a = document.createElement('a');
            const url = URL.createObjectURL(content);
            a.href = url;
            a.download = name + '.docx';
            document.body.appendChild(a);
            a.click();
            setTimeout(function () {
                document.body.removeChild(a);
                window.URL.revokeObjectURL(url);
            }, 0);
        });
    }
    
}

// 解析表格
const parseTable = (table) => {
    let xml = '';
    // 解析表头
    const headerCells = table.querySelectorAll('thead tr th');
    const bodyRows = table.querySelectorAll('tbody tr');
    let headerXml = '<w:tr>';
    for (let i = 0; i < headerCells.length; i++) {
      headerXml += createHeaderCell(headerCells[i].innerText.trim());
    }
    headerXml += '</w:tr>';

    let bodyXml = '';
    for (let i = 0; i < bodyRows.length; i++) { 
        bodyXml += '<w:tr>';
        const bodyCells = bodyRows[i].querySelectorAll('td')
        for (let j = 0; j < bodyCells.length; j++) {
            bodyXml += createBodyCell(bodyCells[j].innerText.trim());
        }
        bodyXml += '</w:tr>';
    }

    xml = `<w:tbl> <w:tblPr> <w:tblStyle w:val="GridTable1Light"/> <w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr> <w:tblGrid> <w:gridCol w:w="2226"/> <w:gridCol w:w="2226"/> <w:gridCol w:w="2226"/> </w:tblGrid> ${headerXml} ${bodyXml}</w:tbl>`

    return xml
}

// 创建表头单元格
const createHeaderCell = (text) => {
    const cell = `
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="0" w:type="auto"/>
          <w:shd w:val="clear" w:color="auto"/>
        </w:tcPr>
        <w:p>
          <w:r>
            <w:rPr>
              <w:b/>
            </w:rPr>
            <w:t>${text}</w:t>
          </w:r>
        </w:p>
      </w:tc>
    `;
    return cell;
  };
  
  // 创建表体单元格
  const createBodyCell = (text) => {
    const cell = `
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="0" w:type="auto"/>
          <w:shd w:val="clear" w:color="auto"/>
        </w:tcPr>
        <w:p>
          <w:r>
            <w:t>${text}</w:t>
          </w:r>
        </w:p>
      </w:tc>
    `;
    return cell;
  };
  

function setContentTypes(zip) {
    // 构建docx所需的xml字符串
    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types
        xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
        <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
        <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
        <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
        <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
        <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
        <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
        <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
        <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    </Types>`;

    zip.file('[Content_Types].xml', contentTypesXml);
}

function setRels(zip) {
    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships
        xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`;

    zip.folder('_rels').file('.rels', relsXml);
}

function setCustomXml(zip) {
    const item1XMLRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships
        xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/>
    </Relationships>`
    const item1XML = `<?xml version="1.0" standalone="no"?>
    <b:Sources
        xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
        xmlns="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" SelectedStyle="\APASixthEditionOfficeOnline.xsl" StyleName="APA" Version="6">
    </b:Sources>`
    const itemProps1XML = `<?xml version="1.0" encoding="UTF-8" standalone="no"?>
    <ds:datastoreItem ds:itemID="{52C8F2C5-85D6-4A2B-9C2B-0C7D136F0F1F}"
        xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">
        <ds:schemaRefs>
            <ds:schemaRef ds:uri="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"/>
        </ds:schemaRefs>
    </ds:datastoreItem>`

    const customXml = zip.folder('customXml')
    customXml.folder('_rels').file('item1.xml.rels', item1XMLRels);
    customXml.file('item1.xml', item1XML);
    customXml.file('itemProps1.xml', itemProps1XML);
    
}

function setDocProps(zip,
    charCount = 2,
    pCount = 1,
    charSpaceCount = 2
) {
    const docProps = zip.folder('docProps');

    docProps.file('app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Properties
        xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
        xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
        <Template>Normal.dotm</Template>
        <TotalTime>81</TotalTime>
        <Pages>3</Pages>
        <Words>105</Words>
        <Characters>${charCount}</Characters>
        <Application>Microsoft Office Word</Application>
        <DocSecurity>0</DocSecurity>
        <Lines>5</Lines>
        <Paragraphs>${pCount}</Paragraphs>
        <ScaleCrop>false</ScaleCrop>
        <HeadingPairs>
            <vt:vector size="2" baseType="variant">
                <vt:variant>
                    <vt:lpstr>Title</vt:lpstr>
                </vt:variant>
                <vt:variant>
                    <vt:i4>1</vt:i4>
                </vt:variant>
            </vt:vector>
        </HeadingPairs>
        <TitlesOfParts>
            <vt:vector size="1" baseType="lpstr">
                <vt:lpstr></vt:lpstr>
            </vt:vector>
        </TitlesOfParts>
        <Company></Company>
        <LinksUpToDate>false</LinksUpToDate>
        <CharactersWithSpaces>${charSpaceCount}</CharactersWithSpaces>
        <SharedDoc>false</SharedDoc>
        <HyperlinksChanged>false</HyperlinksChanged>
        <AppVersion>16.0000</AppVersion>
    </Properties>`);

    docProps.file('core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <cp:coreProperties
      xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
      xmlns:dc="http://purl.org/dc/elements/1.1/"
      xmlns:dcterms="http://purl.org/dc/terms/"
      xmlns:dcmitype="http://purl.org/dc/dcmitype/"
      xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
      <dc:title></dc:title>
      <dc:subject></dc:subject>
      <dc:creator>ExportGPT</dc:creator>
      <cp:keywords></cp:keywords>
      <dc:description></dc:description>
      <cp:lastModifiedBy>ExportGPT</cp:lastModifiedBy>
      <cp:revision>1</cp:revision>
      <dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created>
      <dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:modified>
  </cp:coreProperties>`);
}

function setWord(zip, xml,hyperlinkXML) {
    const wordRelsXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships
        xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
        <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
        <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
        <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
        ${hyperlinkXML}
    </Relationships>`

    const documentXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document
        xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
        xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
        xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
        xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
        xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"
        xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
        xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"
        xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"
        xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"
        xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink"
        xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:oel="http://schemas.microsoft.com/office/2019/extlst"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
        xmlns:v="urn:schemas-microsoft-com:vml"
        xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        xmlns:w10="urn:schemas-microsoft-com:office:word"
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
        xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
        xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
        xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
        xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
        xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"
        xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
        xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
        xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
        xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
        xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
        xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14">
        <w:body>
        ${xml}
            <w:sectPr w:rsidR="009F7532">
                <w:pgSz w:w="11906" w:h="16838"/>
                <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>
                <w:cols w:space="425"/>
                <w:docGrid w:type="lines" w:linePitch="312"/>
            </w:sectPr>
        </w:body>
    </w:document>`

    const fontTableXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:fonts
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
        xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
        xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
        xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
        xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
        xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
        xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh">
        <w:font w:name="Wingdings">
            <w:panose1 w:val="05000000000000000000"/>
            <w:charset w:val="02"/>
            <w:family w:val="auto"/>
            <w:pitch w:val="variable"/>
            <w:sig w:usb0="00000000" w:usb1="10000000" w:usb2="00000000" w:usb3="00000000" w:csb0="80000000" w:csb1="00000000"/>
        </w:font>
        <w:font w:name="Times New Roman">
            <w:panose1 w:val="02020603050405020304"/>
            <w:charset w:val="00"/>
            <w:family w:val="roman"/>
            <w:pitch w:val="variable"/>
            <w:sig w:usb0="E0002EFF" w:usb1="C000785B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
        </w:font>
        <w:font w:name="等线">
            <w:altName w:val="DengXian"/>
            <w:panose1 w:val="02010600030101010101"/>
            <w:charset w:val="86"/>
            <w:family w:val="auto"/>
            <w:pitch w:val="variable"/>
            <w:sig w:usb0="A00002BF" w:usb1="38CF7CFA" w:usb2="00000016" w:usb3="00000000" w:csb0="0004000F" w:csb1="00000000"/>
        </w:font>
        <w:font w:name="等线 Light">
            <w:panose1 w:val="02010600030101010101"/>
            <w:charset w:val="86"/>
            <w:family w:val="auto"/>
            <w:pitch w:val="variable"/>
            <w:sig w:usb0="A00002BF" w:usb1="38CF7CFA" w:usb2="00000016" w:usb3="00000000" w:csb0="0004000F" w:csb1="00000000"/>
        </w:font>
        <w:font w:name="Consolas">
            <w:panose1 w:val="020B0609020204030204"/>
            <w:charset w:val="00"/>
            <w:family w:val="modern"/>
            <w:pitch w:val="fixed"/>
            <w:sig w:usb0="E00006FF" w:usb1="0000FCFF" w:usb2="00000001" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/>
        </w:font>
        <w:font w:name="宋体">
            <w:altName w:val="SimSun"/>
            <w:panose1 w:val="02010600030101010101"/>
            <w:charset w:val="86"/>
            <w:family w:val="auto"/>
            <w:pitch w:val="variable"/>
            <w:sig w:usb0="00000203" w:usb1="288F0000" w:usb2="00000016" w:usb3="00000000" w:csb0="00040001" w:csb1="00000000"/>
        </w:font>
    </w:fonts>`

    const numberingXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:numbering
        xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
        xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
        xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
        xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
        xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"
        xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
        xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"
        xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"
        xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"
        xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink"
        xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:oel="http://schemas.microsoft.com/office/2019/extlst"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
        xmlns:v="urn:schemas-microsoft-com:vml"
        xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        xmlns:w10="urn:schemas-microsoft-com:office:word"
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
        xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
        xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
        xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
        xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
        xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"
        xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
        xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
        xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
        xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
        xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
        xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14">
        <w:abstractNum w:abstractNumId="0" w15:restartNumberingAfterBreak="0">
            <w:nsid w:val="001A0BFA"/>
            <w:multiLevelType w:val="hybridMultilevel"/>
            <w:tmpl w:val="38940A6A"/>
            <w:lvl w:ilvl="0" w:tplc="0409000F">
                <w:start w:val="1"/>
                <w:numFmt w:val="decimal"/>
                <w:lvlText w:val="%1."/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="440" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="1" w:tplc="04090019" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="lowerLetter"/>
                <w:lvlText w:val="%2)"/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="880" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="2" w:tplc="0409001B" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="lowerRoman"/>
                <w:lvlText w:val="%3."/>
                <w:lvlJc w:val="right"/>
                <w:pPr>
                    <w:ind w:left="1320" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="3" w:tplc="0409000F" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="decimal"/>
                <w:lvlText w:val="%4."/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="1760" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="4" w:tplc="04090019" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="lowerLetter"/>
                <w:lvlText w:val="%5)"/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="2200" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="5" w:tplc="0409001B" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="lowerRoman"/>
                <w:lvlText w:val="%6."/>
                <w:lvlJc w:val="right"/>
                <w:pPr>
                    <w:ind w:left="2640" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="6" w:tplc="0409000F" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="decimal"/>
                <w:lvlText w:val="%7."/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3080" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="7" w:tplc="04090019" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="lowerLetter"/>
                <w:lvlText w:val="%8)"/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3520" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="8" w:tplc="0409001B" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="lowerRoman"/>
                <w:lvlText w:val="%9."/>
                <w:lvlJc w:val="right"/>
                <w:pPr>
                    <w:ind w:left="3960" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
        </w:abstractNum>
        <w:abstractNum w:abstractNumId="1" w15:restartNumberingAfterBreak="0">
            <w:nsid w:val="00526A78"/>
            <w:multiLevelType w:val="hybridMultilevel"/>
            <w:tmpl w:val="7208F61E"/>
            <w:lvl w:ilvl="0" w:tplc="04090001">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="360" w:hanging="360"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="1" w:tplc="FFFFFFFF" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="880" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="2" w:tplc="FFFFFFFF" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="1320" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="3" w:tplc="FFFFFFFF" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="1760" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="4" w:tplc="FFFFFFFF" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="2200" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="5" w:tplc="FFFFFFFF" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="2640" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="6" w:tplc="FFFFFFFF" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3080" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="7" w:tplc="FFFFFFFF" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3520" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="8" w:tplc="FFFFFFFF" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3960" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
        </w:abstractNum>
        <w:abstractNum w:abstractNumId="2" w15:restartNumberingAfterBreak="0">
            <w:nsid w:val="1D4E5DBB"/>
            <w:multiLevelType w:val="hybridMultilevel"/>
            <w:tmpl w:val="4D60F5C6"/>
            <w:lvl w:ilvl="0" w:tplc="04090001">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="440" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="1" w:tplc="04090003" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="880" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="2" w:tplc="04090005" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="1320" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="3" w:tplc="04090001" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="1760" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="4" w:tplc="04090003" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="2200" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="5" w:tplc="04090005" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="2640" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="6" w:tplc="04090001" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3080" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="7" w:tplc="04090003" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3520" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="8" w:tplc="04090005" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3960" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
        </w:abstractNum>
        <w:abstractNum w:abstractNumId="3" w15:restartNumberingAfterBreak="0">
            <w:nsid w:val="2AAE0370"/>
            <w:multiLevelType w:val="hybridMultilevel"/>
            <w:tmpl w:val="0990486A"/>
            <w:lvl w:ilvl="0" w:tplc="B26A3C2E">
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val="-"/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="360" w:hanging="360"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="等线" w:eastAsia="等线" w:hAnsi="等线" w:cstheme="minorBidi" w:hint="eastAsia"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="1" w:tplc="04090003" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="880" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="2" w:tplc="04090005" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="1320" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="3" w:tplc="04090001" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="1760" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="4" w:tplc="04090003" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="2200" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="5" w:tplc="04090005" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="2640" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="6" w:tplc="04090001" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3080" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="7" w:tplc="04090003" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3520" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="8" w:tplc="04090005" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3960" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
        </w:abstractNum>
        <w:abstractNum w:abstractNumId="4" w15:restartNumberingAfterBreak="0">
            <w:nsid w:val="2FF806E6"/>
            <w:multiLevelType w:val="hybridMultilevel"/>
            <w:tmpl w:val="6D6C5D12"/>
            <w:lvl w:ilvl="0" w:tplc="90663F8A">
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val="-"/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="360" w:hanging="360"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="等线" w:eastAsia="等线" w:hAnsi="等线" w:cstheme="minorBidi" w:hint="eastAsia"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="1" w:tplc="04090003" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="880" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="2" w:tplc="04090005" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="1320" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="3" w:tplc="04090001" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="1760" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="4" w:tplc="04090003" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="2200" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="5" w:tplc="04090005" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="2640" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="6" w:tplc="04090001" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3080" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="7" w:tplc="04090003" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3520" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
            <w:lvl w:ilvl="8" w:tplc="04090005" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="bullet"/>
                <w:lvlText w:val=""/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3960" w:hanging="440"/>
                </w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>
                </w:rPr>
            </w:lvl>
        </w:abstractNum>
        <w:abstractNum w:abstractNumId="5" w15:restartNumberingAfterBreak="0">
            <w:nsid w:val="6D52303C"/>
            <w:multiLevelType w:val="hybridMultilevel"/>
            <w:tmpl w:val="55F8986A"/>
            <w:lvl w:ilvl="0" w:tplc="0409000F">
                <w:start w:val="1"/>
                <w:numFmt w:val="decimal"/>
                <w:lvlText w:val="%1."/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="440" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="1" w:tplc="04090019" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="lowerLetter"/>
                <w:lvlText w:val="%2)"/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="880" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="2" w:tplc="0409001B" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="lowerRoman"/>
                <w:lvlText w:val="%3."/>
                <w:lvlJc w:val="right"/>
                <w:pPr>
                    <w:ind w:left="1320" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="3" w:tplc="0409000F" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="decimal"/>
                <w:lvlText w:val="%4."/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="1760" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="4" w:tplc="04090019" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="lowerLetter"/>
                <w:lvlText w:val="%5)"/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="2200" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="5" w:tplc="0409001B" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="lowerRoman"/>
                <w:lvlText w:val="%6."/>
                <w:lvlJc w:val="right"/>
                <w:pPr>
                    <w:ind w:left="2640" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="6" w:tplc="0409000F" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="decimal"/>
                <w:lvlText w:val="%7."/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3080" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="7" w:tplc="04090019" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="lowerLetter"/>
                <w:lvlText w:val="%8)"/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="3520" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
            <w:lvl w:ilvl="8" w:tplc="0409001B" w:tentative="1">
                <w:start w:val="1"/>
                <w:numFmt w:val="lowerRoman"/>
                <w:lvlText w:val="%9."/>
                <w:lvlJc w:val="right"/>
                <w:pPr>
                    <w:ind w:left="3960" w:hanging="440"/>
                </w:pPr>
            </w:lvl>
        </w:abstractNum>
        <w:num w:numId="1" w16cid:durableId="1197740206">
            <w:abstractNumId w:val="2"/>
        </w:num>
        <w:num w:numId="2" w16cid:durableId="1409569620">
            <w:abstractNumId w:val="0"/>
        </w:num>
        <w:num w:numId="3" w16cid:durableId="1548879684">
            <w:abstractNumId w:val="3"/>
        </w:num>
        <w:num w:numId="4" w16cid:durableId="1014529958">
            <w:abstractNumId w:val="4"/>
        </w:num>
        <w:num w:numId="5" w16cid:durableId="1206602628">
            <w:abstractNumId w:val="1"/>
        </w:num>
        <w:num w:numId="6" w16cid:durableId="1019888690">
            <w:abstractNumId w:val="5"/>
        </w:num>
    </w:numbering>`

    const settingsXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:settings
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
        xmlns:v="urn:schemas-microsoft-com:vml"
        xmlns:w10="urn:schemas-microsoft-com:office:word"
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
        xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
        xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
        xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
        xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
        xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
        xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
        xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh">
        <w:zoom w:percent="100"/>
        <w:bordersDoNotSurroundHeader/>
        <w:bordersDoNotSurroundFooter/>
        <w:proofState w:spelling="clean" w:grammar="clean"/>
        <w:defaultTabStop w:val="420"/>
        <w:drawingGridVerticalSpacing w:val="156"/>
        <w:displayHorizontalDrawingGridEvery w:val="0"/>
        <w:displayVerticalDrawingGridEvery w:val="2"/>
        <w:characterSpacingControl w:val="compressPunctuation"/>
        <w:compat>
            <w:spaceForUL/>
            <w:balanceSingleByteDoubleByteWidth/>
            <w:doNotLeaveBackslashAlone/>
            <w:ulTrailSpace/>
            <w:doNotExpandShiftReturn/>
            <w:adjustLineHeightInTable/>
            <w:useFELayout/>
            <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
            <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
            <w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
            <w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
            <w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
            <w:compatSetting w:name="useWord2013TrackBottomHyphenation" w:uri="http://schemas.microsoft.com/office/word" w:val="0"/>
        </w:compat>
        <w:rsids>
            <w:rsidRoot w:val="00D561EF"/>
            <w:rsid w:val="00145D17"/>
            <w:rsid w:val="00432CFF"/>
            <w:rsid w:val="00776EE0"/>
            <w:rsid w:val="007F6D62"/>
            <w:rsid w:val="0081448F"/>
            <w:rsid w:val="008F028D"/>
            <w:rsid w:val="009E133E"/>
            <w:rsid w:val="00BC210D"/>
            <w:rsid w:val="00CA0062"/>
            <w:rsid w:val="00D561EF"/>
        </w:rsids>
        <m:mathPr>
            <m:mathFont m:val="Cambria Math"/>
            <m:brkBin m:val="before"/>
            <m:brkBinSub m:val="--"/>
            <m:smallFrac m:val="0"/>
            <m:dispDef/>
            <m:lMargin m:val="0"/>
            <m:rMargin m:val="0"/>
            <m:defJc m:val="centerGroup"/>
            <m:wrapIndent m:val="1440"/>
            <m:intLim m:val="subSup"/>
            <m:naryLim m:val="undOvr"/>
        </m:mathPr>
        <w:themeFontLang w:val="en-US" w:eastAsia="zh-CN"/>
        <w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>
        <w:shapeDefaults>
            <o:shapedefaults v:ext="edit" spidmax="1026"/>
            <o:shapelayout v:ext="edit">
                <o:idmap v:ext="edit" data="1"/>
            </o:shapelayout>
        </w:shapeDefaults>
        <w:decimalSymbol w:val="."/>
        <w:listSeparator w:val=","/>
        <w14:docId w14:val="0F86C403"/>
        <w15:chartTrackingRefBased/>
        <w15:docId w15:val="{459ECC98-089E-4F6E-BFC2-4C8913AFEFCE}"/>
    </w:settings>`

    const Normal = `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
        <w:widowControl w:val="0"/>
        <w:jc w:val="both"/>
    </w:pPr>
</w:style>`

    const H1StyleXML = `<w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="Heading1Char"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:rsid w:val="00EE499C"/>
    <w:pPr>
        <w:keepNext/>
        <w:keepLines/>
        <w:spacing w:before="340" w:after="330" w:line="578" w:lineRule="auto"/>
        <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
        <w:b/>
        <w:bCs/>
        <w:kern w:val="44"/>
        <w:sz w:val="44"/>
        <w:szCs w:val="44"/>
    </w:rPr>
</w:style>`
    const H2StyleXML = `<w:style w:type="paragraph" w:styleId="Heading2">
<w:name w:val="heading 2"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Normal"/>
<w:link w:val="Heading2Char"/>
<w:uiPriority w:val="9"/>
<w:unhideWhenUsed/>
<w:qFormat/>
<w:rsid w:val="00EE499C"/>
<w:pPr>
    <w:keepNext/>
    <w:keepLines/>
    <w:spacing w:before="260" w:after="260" w:line="416" w:lineRule="auto"/>
    <w:outlineLvl w:val="1"/>
</w:pPr>
<w:rPr>
    <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
    <w:b/>
    <w:bCs/>
    <w:sz w:val="32"/>
    <w:szCs w:val="32"/>
</w:rPr>
</w:style>`;

    const H3StyleXML = `<w:style w:type="paragraph" w:styleId="Heading3">
<w:name w:val="heading 3"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Normal"/>
<w:link w:val="Heading3Char"/>
<w:uiPriority w:val="9"/>
<w:unhideWhenUsed/>
<w:qFormat/>
<w:rsid w:val="00EE499C"/>
<w:pPr>
    <w:keepNext/>
    <w:keepLines/>
    <w:spacing w:before="260" w:after="260" w:line="416" w:lineRule="auto"/>
    <w:outlineLvl w:val="2"/>
</w:pPr>
<w:rPr>
    <w:b/>
    <w:bCs/>
    <w:sz w:val="32"/>
    <w:szCs w:val="32"/>
</w:rPr>
</w:style>`;

    const H4StyleXML = `<w:style w:type="paragraph" w:styleId="Heading4">
<w:name w:val="heading 4"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Normal"/>
<w:link w:val="Heading4Char"/>
<w:uiPriority w:val="9"/>
<w:unhideWhenUsed/>
<w:qFormat/>
<w:rsid w:val="00EE499C"/>
<w:pPr>
    <w:keepNext/>
    <w:keepLines/>
    <w:spacing w:before="280" w:after="290" w:line="376" w:lineRule="auto"/>
    <w:outlineLvl w:val="3"/>
</w:pPr>
<w:rPr>
    <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
    <w:b/>
    <w:bCs/>
    <w:sz w:val="28"/>
    <w:szCs w:val="28"/>
</w:rPr>
</w:style>`;

    const H5StyleXML = `<w:style w:type="paragraph" w:styleId="Heading5">
<w:name w:val="heading 5"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Normal"/>
<w:link w:val="Heading5Char"/>
<w:uiPriority w:val="9"/>
<w:unhideWhenUsed/>
<w:qFormat/>
<w:rsid w:val="00EE499C"/>
<w:pPr>
    <w:keepNext/>
    <w:keepLines/>
    <w:spacing w:before="280" w:after="290" w:line="376" w:lineRule="auto"/>
    <w:outlineLvl w:val="4"/>
</w:pPr>
<w:rPr>
    <w:b/>
    <w:bCs/>
    <w:sz w:val="28"/>
    <w:szCs w:val="28"/>
</w:rPr>
</w:style>`;

    const H6StyleXML = `<w:style w:type="paragraph" w:styleId="Heading6">
<w:name w:val="heading 6"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Normal"/>
<w:link w:val="Heading6Char"/>
<w:uiPriority w:val="9"/>
<w:unhideWhenUsed/>
<w:qFormat/>
<w:rsid w:val="00EE499C"/>
<w:pPr>
    <w:keepNext/>
    <w:keepLines/>
    <w:spacing w:before="240" w:after="64" w:line="320" w:lineRule="auto"/>
    <w:outlineLvl w:val="5"/>
</w:pPr>
<w:rPr>
    <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
    <w:b/>
    <w:bCs/>
    <w:sz w:val="24"/>
    <w:szCs w:val="24"/>
</w:rPr>
</w:style>`;
    const Heading1Char = ` <w:style w:type="character" w:customStyle="1" w:styleId="Heading1Char">
<w:name w:val="Heading 1 Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Heading1"/>
<w:uiPriority w:val="9"/>
<w:rsid w:val="0079231C"/>
<w:rPr>
    <w:b/>
    <w:bCs/>
    <w:kern w:val="44"/>
    <w:sz w:val="44"/>
    <w:szCs w:val="44"/>
</w:rPr>
</w:style>`;
    const Heading2Char = `<w:style w:type="character" w:customStyle="1" w:styleId="Heading2Char">
<w:name w:val="Heading 2 Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Heading2"/>
<w:uiPriority w:val="9"/>
<w:rsid w:val="00EE499C"/>
<w:rPr>
    <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
    <w:b/>
    <w:bCs/>
    <w:sz w:val="32"/>
    <w:szCs w:val="32"/>
</w:rPr>
</w:style>`;
    const Heading3Char = `<w:style w:type="character" w:customStyle="1" w:styleId="Heading3Char">
<w:name w:val="Heading 3 Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Heading3"/>
<w:uiPriority w:val="9"/>
<w:rsid w:val="00EE499C"/>
<w:rPr>
    <w:b/>
    <w:bCs/>
    <w:sz w:val="32"/>
    <w:szCs w:val="32"/>
</w:rPr>
</w:style>`;
    const Heading4Char = `<w:style w:type="character" w:customStyle="1" w:styleId="Heading4Char">
<w:name w:val="Heading 4 Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Heading4"/>
<w:uiPriority w:val="9"/>
<w:rsid w:val="00EE499C"/>
<w:rPr>
    <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
    <w:b/>
    <w:bCs/>
    <w:sz w:val="28"/>
    <w:szCs w:val="28"/>
</w:rPr>
</w:style>`;
    const Heading5Char = ` <w:style w:type="character" w:customStyle="1" w:styleId="Heading5Char">
<w:name w:val="Heading 5 Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Heading5"/>
<w:uiPriority w:val="9"/>
<w:rsid w:val="00EE499C"/>
<w:rPr>
    <w:b/>
    <w:bCs/>
    <w:sz w:val="28"/>
    <w:szCs w:val="28"/>
</w:rPr>
</w:style>`;
    const Heading6Char = `<w:style w:type="character" w:customStyle="1" w:styleId="Heading6Char">
<w:name w:val="Heading 6 Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Heading6"/>
<w:uiPriority w:val="9"/>
<w:rsid w:val="00EE499C"/>
<w:rPr>
    <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
    <w:b/>
    <w:bCs/>
    <w:sz w:val="24"/>
    <w:szCs w:val="24"/>
</w:rPr>
</w:style>`;



    const Title = `<w:style w:type="paragraph" w:styleId="Title">
<w:name w:val="Title"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Normal"/>
<w:link w:val="TitleChar"/>
<w:uiPriority w:val="10"/>
<w:qFormat/>
<w:rsid w:val="00EE499C"/>
<w:pPr>
    <w:spacing w:before="240" w:after="60"/>
    <w:jc w:val="center"/>
    <w:outlineLvl w:val="0"/>
</w:pPr>
<w:rPr>
    <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
    <w:b/>
    <w:bCs/>
    <w:sz w:val="32"/>
    <w:szCs w:val="32"/>
</w:rPr>
</w:style>`
    const TitleChar = `<w:style w:type="character" w:customStyle="1" w:styleId="TitleChar">
<w:name w:val="Title Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Title"/>
<w:uiPriority w:val="10"/>
<w:rsid w:val="00EE499C"/>
<w:rPr>
    <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
    <w:b/>
    <w:bCs/>
    <w:sz w:val="32"/>
    <w:szCs w:val="32"/>
</w:rPr>
</w:style>`
    const Subtitle = `<w:style w:type="paragraph" w:styleId="Subtitle">
<w:name w:val="Subtitle"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Normal"/>
<w:link w:val="SubtitleChar"/>
<w:uiPriority w:val="11"/>
<w:qFormat/>
<w:rsid w:val="00EE499C"/>
<w:pPr>
    <w:spacing w:before="240" w:after="60" w:line="312" w:lineRule="auto"/>
    <w:jc w:val="center"/>
    <w:outlineLvl w:val="1"/>
</w:pPr>
<w:rPr>
    <w:b/>
    <w:bCs/>
    <w:kern w:val="28"/>
    <w:sz w:val="32"/>
    <w:szCs w:val="32"/>
</w:rPr>
</w:style>`
    const SubtitleChar = `<w:style w:type="character" w:customStyle="1" w:styleId="SubtitleChar">
<w:name w:val="Subtitle Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Subtitle"/>
<w:uiPriority w:val="11"/>
<w:rsid w:val="00EE499C"/>
<w:rPr>
    <w:b/>
    <w:bCs/>
    <w:kern w:val="28"/>
    <w:sz w:val="32"/>
    <w:szCs w:val="32"/>
</w:rPr>
</w:style>`

    const Emphasis = `<w:style w:type="character" w:styleId="Emphasis">
<w:name w:val="Emphasis"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:uiPriority w:val="20"/>
<w:qFormat/>
<w:rsid w:val="00FF4E92"/>
<w:rPr>
    <w:i/>
    <w:iCs/>
</w:rPr>
</w:style>`
    const Strong = `<w:style w:type="character" w:styleId="Strong">
<w:name w:val="Strong"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:uiPriority w:val="22"/>
<w:qFormat/>
<w:rsid w:val="00FF4E92"/>
<w:rPr>
    <w:b/>
    <w:bCs/>
</w:rPr>
</w:style>`
    const Quote = `<w:style w:type="paragraph" w:styleId="Quote">
<w:name w:val="Quote"/>
<w:basedOn w:val="Normal"/>
<w:next w:val="Normal"/>
<w:link w:val="QuoteChar"/>
<w:uiPriority w:val="29"/>
<w:qFormat/>
<w:rsid w:val="00FF4E92"/>
<w:pPr>
    <w:spacing w:before="200" w:after="160"/>
    <w:ind w:left="864" w:right="864"/>
    <w:jc w:val="center"/>
</w:pPr>
<w:rPr>
    <w:i/>
    <w:iCs/>
    <w:color w:val="404040" w:themeColor="text1" w:themeTint="BF"/>
</w:rPr>
</w:style>`
    const QuoteChar = `<w:style w:type="character" w:customStyle="1" w:styleId="QuoteChar">
<w:name w:val="Quote Char"/>
<w:basedOn w:val="DefaultParagraphFont"/>
<w:link w:val="Quote"/>
<w:uiPriority w:val="29"/>
<w:rsid w:val="00FF4E92"/>
<w:rPr>
    <w:i/>
    <w:iCs/>
    <w:color w:val="404040" w:themeColor="text1" w:themeTint="BF"/>
</w:rPr>
</w:style>`
    const Hyperlink = `<w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="99"/>
    <w:unhideWhenUsed/>
    <w:rsid w:val="00776EE0"/>
    <w:rPr>
        <w:color w:val="0563C1" w:themeColor="hyperlink"/>
        <w:u w:val="single"/>
    </w:rPr>
</w:style>`
    const GridTable1Light = `<w:style w:type="table" w:styleId="GridTable1Light">
    <w:name w:val="Grid Table 1 Light"/>
    <w:basedOn w:val="TableNormal"/>
    <w:uiPriority w:val="46"/>
    <w:rsid w:val="009E133E"/>
    <w:tblPr>
        <w:tblStyleRowBandSize w:val="1"/>
        <w:tblStyleColBandSize w:val="1"/>
        <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="999999" w:themeColor="text1" w:themeTint="66"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="999999" w:themeColor="text1" w:themeTint="66"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="999999" w:themeColor="text1" w:themeTint="66"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="999999" w:themeColor="text1" w:themeTint="66"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="999999" w:themeColor="text1" w:themeTint="66"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="999999" w:themeColor="text1" w:themeTint="66"/>
        </w:tblBorders>
    </w:tblPr>
    <w:tblStylePr w:type="firstRow">
        <w:rPr>
            <w:b/>
            <w:bCs/>
        </w:rPr>
        <w:tblPr/>
        <w:tcPr>
            <w:tcBorders>
                <w:bottom w:val="single" w:sz="12" w:space="0" w:color="666666" w:themeColor="text1" w:themeTint="99"/>
            </w:tcBorders>
        </w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="lastRow">
        <w:rPr>
            <w:b/>
            <w:bCs/>
        </w:rPr>
        <w:tblPr/>
        <w:tcPr>
            <w:tcBorders>
                <w:top w:val="double" w:sz="2" w:space="0" w:color="666666" w:themeColor="text1" w:themeTint="99"/>
            </w:tcBorders>
        </w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="firstCol">
        <w:rPr>
            <w:b/>
            <w:bCs/>
        </w:rPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="lastCol">
        <w:rPr>
            <w:b/>
            <w:bCs/>
        </w:rPr>
    </w:tblStylePr>
</w:style>`

    const DefaultParagraphFont = ` <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
<w:name w:val="Default Paragraph Font"/>
<w:uiPriority w:val="1"/>
<w:semiHidden/>
<w:unhideWhenUsed/>
</w:style>`


    const TableNormal = `<w:style w:type="table" w:default="1" w:styleId="TableNormal">
<w:name w:val="Normal Table"/>
<w:uiPriority w:val="99"/>
<w:semiHidden/>
<w:unhideWhenUsed/>
<w:tblPr>
    <w:tblInd w:w="0" w:type="dxa"/>
    <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/>
        <w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/>
        <w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar>
</w:tblPr>
</w:style>`

    const NoList = `<w:style w:type="numbering" w:default="1" w:styleId="NoList">
<w:name w:val="No List"/>
<w:uiPriority w:val="99"/>
<w:semiHidden/>
<w:unhideWhenUsed/>
</w:style>`
    const ListParagraph = `<w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
    <w:basedOn w:val="Normal"/>
    <w:uiPriority w:val="34"/>
    <w:qFormat/>
    <w:rsid w:val="008F028D"/>
    <w:pPr>
        <w:ind w:firstLineChars="200" w:firstLine="420"/>
    </w:pPr>
</w:style>`


    const stylesXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:styles
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
        xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
        xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
        xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
        xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
        xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
        xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh">
        <w:docDefaults>
            <w:rPrDefault>
                <w:rPr>
                    <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
                    <w:kern w:val="2"/>
                    <w:sz w:val="21"/>
                    <w:szCs w:val="22"/>
                    <w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>
                </w:rPr>
            </w:rPrDefault>
            <w:pPrDefault/>
        </w:docDefaults>
        <w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="376">
            <w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>
            <w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>
            <w:lsdException w:name="heading 2" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
            <w:lsdException w:name="heading 3" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
            <w:lsdException w:name="heading 4" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
            <w:lsdException w:name="heading 5" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
            <w:lsdException w:name="heading 6" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
            <w:lsdException w:name="heading 7" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
            <w:lsdException w:name="heading 8" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
            <w:lsdException w:name="heading 9" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
            <w:lsdException w:name="index 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="index 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="index 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="index 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="index 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="index 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="index 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="index 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="index 9" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="toc 1" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="toc 2" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="toc 3" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="toc 4" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="toc 5" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="toc 6" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="toc 7" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="toc 8" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="toc 9" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Normal Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="footnote text" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="annotation text" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="header" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="footer" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="index heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="caption" w:semiHidden="1" w:uiPriority="35" w:unhideWhenUsed="1" w:qFormat="1"/>
            <w:lsdException w:name="table of figures" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="envelope address" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="envelope return" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="footnote reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="annotation reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="line number" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="page number" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="endnote reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="endnote text" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="table of authorities" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="macro" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="toa heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Bullet" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Number" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Bullet 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Bullet 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Bullet 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Bullet 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Number 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Number 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Number 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Number 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Title" w:uiPriority="10" w:qFormat="1"/>
            <w:lsdException w:name="Closing" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Signature" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Default Paragraph Font" w:semiHidden="1" w:uiPriority="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Body Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Body Text Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Continue" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Continue 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Continue 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Continue 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="List Continue 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Message Header" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Subtitle" w:uiPriority="11" w:qFormat="1"/>
            <w:lsdException w:name="Salutation" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Date" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Body Text First Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Body Text First Indent 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Note Heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Body Text 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Body Text 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Body Text Indent 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Body Text Indent 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Block Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Hyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="FollowedHyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Strong" w:uiPriority="22" w:qFormat="1"/>
            <w:lsdException w:name="Emphasis" w:uiPriority="20" w:qFormat="1"/>
            <w:lsdException w:name="Document Map" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Plain Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="E-mail Signature" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="HTML Top of Form" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="HTML Bottom of Form" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Normal (Web)" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="HTML Acronym" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="HTML Address" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="HTML Cite" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="HTML Code" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="HTML Definition" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="HTML Keyboard" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="HTML Preformatted" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="HTML Sample" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="HTML Typewriter" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="HTML Variable" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Normal Table" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="annotation subject" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="No List" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Outline List 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Outline List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Outline List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Simple 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Simple 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Simple 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Classic 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Classic 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Classic 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Classic 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Colorful 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Colorful 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Colorful 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Columns 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Columns 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Columns 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Columns 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Columns 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Grid 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Grid 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Grid 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Grid 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Grid 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Grid 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Grid 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Grid 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table List 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table List 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table List 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table List 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table List 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table List 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table 3D effects 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table 3D effects 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table 3D effects 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Contemporary" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Elegant" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Professional" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Subtle 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Subtle 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Web 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Web 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Web 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Balloon Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Table Grid" w:uiPriority="39"/>
            <w:lsdException w:name="Table Theme" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Placeholder Text" w:semiHidden="1"/>
            <w:lsdException w:name="No Spacing" w:uiPriority="1" w:qFormat="1"/>
            <w:lsdException w:name="Light Shading" w:uiPriority="60"/>
            <w:lsdException w:name="Light List" w:uiPriority="61"/>
            <w:lsdException w:name="Light Grid" w:uiPriority="62"/>
            <w:lsdException w:name="Medium Shading 1" w:uiPriority="63"/>
            <w:lsdException w:name="Medium Shading 2" w:uiPriority="64"/>
            <w:lsdException w:name="Medium List 1" w:uiPriority="65"/>
            <w:lsdException w:name="Medium List 2" w:uiPriority="66"/>
            <w:lsdException w:name="Medium Grid 1" w:uiPriority="67"/>
            <w:lsdException w:name="Medium Grid 2" w:uiPriority="68"/>
            <w:lsdException w:name="Medium Grid 3" w:uiPriority="69"/>
            <w:lsdException w:name="Dark List" w:uiPriority="70"/>
            <w:lsdException w:name="Colorful Shading" w:uiPriority="71"/>
            <w:lsdException w:name="Colorful List" w:uiPriority="72"/>
            <w:lsdException w:name="Colorful Grid" w:uiPriority="73"/>
            <w:lsdException w:name="Light Shading Accent 1" w:uiPriority="60"/>
            <w:lsdException w:name="Light List Accent 1" w:uiPriority="61"/>
            <w:lsdException w:name="Light Grid Accent 1" w:uiPriority="62"/>
            <w:lsdException w:name="Medium Shading 1 Accent 1" w:uiPriority="63"/>
            <w:lsdException w:name="Medium Shading 2 Accent 1" w:uiPriority="64"/>
            <w:lsdException w:name="Medium List 1 Accent 1" w:uiPriority="65"/>
            <w:lsdException w:name="Revision" w:semiHidden="1"/>
            <w:lsdException w:name="List Paragraph" w:uiPriority="34" w:qFormat="1"/>
            <w:lsdException w:name="Quote" w:uiPriority="29" w:qFormat="1"/>
            <w:lsdException w:name="Intense Quote" w:uiPriority="30" w:qFormat="1"/>
            <w:lsdException w:name="Medium List 2 Accent 1" w:uiPriority="66"/>
            <w:lsdException w:name="Medium Grid 1 Accent 1" w:uiPriority="67"/>
            <w:lsdException w:name="Medium Grid 2 Accent 1" w:uiPriority="68"/>
            <w:lsdException w:name="Medium Grid 3 Accent 1" w:uiPriority="69"/>
            <w:lsdException w:name="Dark List Accent 1" w:uiPriority="70"/>
            <w:lsdException w:name="Colorful Shading Accent 1" w:uiPriority="71"/>
            <w:lsdException w:name="Colorful List Accent 1" w:uiPriority="72"/>
            <w:lsdException w:name="Colorful Grid Accent 1" w:uiPriority="73"/>
            <w:lsdException w:name="Light Shading Accent 2" w:uiPriority="60"/>
            <w:lsdException w:name="Light List Accent 2" w:uiPriority="61"/>
            <w:lsdException w:name="Light Grid Accent 2" w:uiPriority="62"/>
            <w:lsdException w:name="Medium Shading 1 Accent 2" w:uiPriority="63"/>
            <w:lsdException w:name="Medium Shading 2 Accent 2" w:uiPriority="64"/>
            <w:lsdException w:name="Medium List 1 Accent 2" w:uiPriority="65"/>
            <w:lsdException w:name="Medium List 2 Accent 2" w:uiPriority="66"/>
            <w:lsdException w:name="Medium Grid 1 Accent 2" w:uiPriority="67"/>
            <w:lsdException w:name="Medium Grid 2 Accent 2" w:uiPriority="68"/>
            <w:lsdException w:name="Medium Grid 3 Accent 2" w:uiPriority="69"/>
            <w:lsdException w:name="Dark List Accent 2" w:uiPriority="70"/>
            <w:lsdException w:name="Colorful Shading Accent 2" w:uiPriority="71"/>
            <w:lsdException w:name="Colorful List Accent 2" w:uiPriority="72"/>
            <w:lsdException w:name="Colorful Grid Accent 2" w:uiPriority="73"/>
            <w:lsdException w:name="Light Shading Accent 3" w:uiPriority="60"/>
            <w:lsdException w:name="Light List Accent 3" w:uiPriority="61"/>
            <w:lsdException w:name="Light Grid Accent 3" w:uiPriority="62"/>
            <w:lsdException w:name="Medium Shading 1 Accent 3" w:uiPriority="63"/>
            <w:lsdException w:name="Medium Shading 2 Accent 3" w:uiPriority="64"/>
            <w:lsdException w:name="Medium List 1 Accent 3" w:uiPriority="65"/>
            <w:lsdException w:name="Medium List 2 Accent 3" w:uiPriority="66"/>
            <w:lsdException w:name="Medium Grid 1 Accent 3" w:uiPriority="67"/>
            <w:lsdException w:name="Medium Grid 2 Accent 3" w:uiPriority="68"/>
            <w:lsdException w:name="Medium Grid 3 Accent 3" w:uiPriority="69"/>
            <w:lsdException w:name="Dark List Accent 3" w:uiPriority="70"/>
            <w:lsdException w:name="Colorful Shading Accent 3" w:uiPriority="71"/>
            <w:lsdException w:name="Colorful List Accent 3" w:uiPriority="72"/>
            <w:lsdException w:name="Colorful Grid Accent 3" w:uiPriority="73"/>
            <w:lsdException w:name="Light Shading Accent 4" w:uiPriority="60"/>
            <w:lsdException w:name="Light List Accent 4" w:uiPriority="61"/>
            <w:lsdException w:name="Light Grid Accent 4" w:uiPriority="62"/>
            <w:lsdException w:name="Medium Shading 1 Accent 4" w:uiPriority="63"/>
            <w:lsdException w:name="Medium Shading 2 Accent 4" w:uiPriority="64"/>
            <w:lsdException w:name="Medium List 1 Accent 4" w:uiPriority="65"/>
            <w:lsdException w:name="Medium List 2 Accent 4" w:uiPriority="66"/>
            <w:lsdException w:name="Medium Grid 1 Accent 4" w:uiPriority="67"/>
            <w:lsdException w:name="Medium Grid 2 Accent 4" w:uiPriority="68"/>
            <w:lsdException w:name="Medium Grid 3 Accent 4" w:uiPriority="69"/>
            <w:lsdException w:name="Dark List Accent 4" w:uiPriority="70"/>
            <w:lsdException w:name="Colorful Shading Accent 4" w:uiPriority="71"/>
            <w:lsdException w:name="Colorful List Accent 4" w:uiPriority="72"/>
            <w:lsdException w:name="Colorful Grid Accent 4" w:uiPriority="73"/>
            <w:lsdException w:name="Light Shading Accent 5" w:uiPriority="60"/>
            <w:lsdException w:name="Light List Accent 5" w:uiPriority="61"/>
            <w:lsdException w:name="Light Grid Accent 5" w:uiPriority="62"/>
            <w:lsdException w:name="Medium Shading 1 Accent 5" w:uiPriority="63"/>
            <w:lsdException w:name="Medium Shading 2 Accent 5" w:uiPriority="64"/>
            <w:lsdException w:name="Medium List 1 Accent 5" w:uiPriority="65"/>
            <w:lsdException w:name="Medium List 2 Accent 5" w:uiPriority="66"/>
            <w:lsdException w:name="Medium Grid 1 Accent 5" w:uiPriority="67"/>
            <w:lsdException w:name="Medium Grid 2 Accent 5" w:uiPriority="68"/>
            <w:lsdException w:name="Medium Grid 3 Accent 5" w:uiPriority="69"/>
            <w:lsdException w:name="Dark List Accent 5" w:uiPriority="70"/>
            <w:lsdException w:name="Colorful Shading Accent 5" w:uiPriority="71"/>
            <w:lsdException w:name="Colorful List Accent 5" w:uiPriority="72"/>
            <w:lsdException w:name="Colorful Grid Accent 5" w:uiPriority="73"/>
            <w:lsdException w:name="Light Shading Accent 6" w:uiPriority="60"/>
            <w:lsdException w:name="Light List Accent 6" w:uiPriority="61"/>
            <w:lsdException w:name="Light Grid Accent 6" w:uiPriority="62"/>
            <w:lsdException w:name="Medium Shading 1 Accent 6" w:uiPriority="63"/>
            <w:lsdException w:name="Medium Shading 2 Accent 6" w:uiPriority="64"/>
            <w:lsdException w:name="Medium List 1 Accent 6" w:uiPriority="65"/>
            <w:lsdException w:name="Medium List 2 Accent 6" w:uiPriority="66"/>
            <w:lsdException w:name="Medium Grid 1 Accent 6" w:uiPriority="67"/>
            <w:lsdException w:name="Medium Grid 2 Accent 6" w:uiPriority="68"/>
            <w:lsdException w:name="Medium Grid 3 Accent 6" w:uiPriority="69"/>
            <w:lsdException w:name="Dark List Accent 6" w:uiPriority="70"/>
            <w:lsdException w:name="Colorful Shading Accent 6" w:uiPriority="71"/>
            <w:lsdException w:name="Colorful List Accent 6" w:uiPriority="72"/>
            <w:lsdException w:name="Colorful Grid Accent 6" w:uiPriority="73"/>
            <w:lsdException w:name="Subtle Emphasis" w:uiPriority="19" w:qFormat="1"/>
            <w:lsdException w:name="Intense Emphasis" w:uiPriority="21" w:qFormat="1"/>
            <w:lsdException w:name="Subtle Reference" w:uiPriority="31" w:qFormat="1"/>
            <w:lsdException w:name="Intense Reference" w:uiPriority="32" w:qFormat="1"/>
            <w:lsdException w:name="Book Title" w:uiPriority="33" w:qFormat="1"/>
            <w:lsdException w:name="Bibliography" w:semiHidden="1" w:uiPriority="37" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="TOC Heading" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1" w:qFormat="1"/>
            <w:lsdException w:name="Plain Table 1" w:uiPriority="41"/>
            <w:lsdException w:name="Plain Table 2" w:uiPriority="42"/>
            <w:lsdException w:name="Plain Table 3" w:uiPriority="43"/>
            <w:lsdException w:name="Plain Table 4" w:uiPriority="44"/>
            <w:lsdException w:name="Plain Table 5" w:uiPriority="45"/>
            <w:lsdException w:name="Grid Table Light" w:uiPriority="40"/>
            <w:lsdException w:name="Grid Table 1 Light" w:uiPriority="46"/>
            <w:lsdException w:name="Grid Table 2" w:uiPriority="47"/>
            <w:lsdException w:name="Grid Table 3" w:uiPriority="48"/>
            <w:lsdException w:name="Grid Table 4" w:uiPriority="49"/>
            <w:lsdException w:name="Grid Table 5 Dark" w:uiPriority="50"/>
            <w:lsdException w:name="Grid Table 6 Colorful" w:uiPriority="51"/>
            <w:lsdException w:name="Grid Table 7 Colorful" w:uiPriority="52"/>
            <w:lsdException w:name="Grid Table 1 Light Accent 1" w:uiPriority="46"/>
            <w:lsdException w:name="Grid Table 2 Accent 1" w:uiPriority="47"/>
            <w:lsdException w:name="Grid Table 3 Accent 1" w:uiPriority="48"/>
            <w:lsdException w:name="Grid Table 4 Accent 1" w:uiPriority="49"/>
            <w:lsdException w:name="Grid Table 5 Dark Accent 1" w:uiPriority="50"/>
            <w:lsdException w:name="Grid Table 6 Colorful Accent 1" w:uiPriority="51"/>
            <w:lsdException w:name="Grid Table 7 Colorful Accent 1" w:uiPriority="52"/>
            <w:lsdException w:name="Grid Table 1 Light Accent 2" w:uiPriority="46"/>
            <w:lsdException w:name="Grid Table 2 Accent 2" w:uiPriority="47"/>
            <w:lsdException w:name="Grid Table 3 Accent 2" w:uiPriority="48"/>
            <w:lsdException w:name="Grid Table 4 Accent 2" w:uiPriority="49"/>
            <w:lsdException w:name="Grid Table 5 Dark Accent 2" w:uiPriority="50"/>
            <w:lsdException w:name="Grid Table 6 Colorful Accent 2" w:uiPriority="51"/>
            <w:lsdException w:name="Grid Table 7 Colorful Accent 2" w:uiPriority="52"/>
            <w:lsdException w:name="Grid Table 1 Light Accent 3" w:uiPriority="46"/>
            <w:lsdException w:name="Grid Table 2 Accent 3" w:uiPriority="47"/>
            <w:lsdException w:name="Grid Table 3 Accent 3" w:uiPriority="48"/>
            <w:lsdException w:name="Grid Table 4 Accent 3" w:uiPriority="49"/>
            <w:lsdException w:name="Grid Table 5 Dark Accent 3" w:uiPriority="50"/>
            <w:lsdException w:name="Grid Table 6 Colorful Accent 3" w:uiPriority="51"/>
            <w:lsdException w:name="Grid Table 7 Colorful Accent 3" w:uiPriority="52"/>
            <w:lsdException w:name="Grid Table 1 Light Accent 4" w:uiPriority="46"/>
            <w:lsdException w:name="Grid Table 2 Accent 4" w:uiPriority="47"/>
            <w:lsdException w:name="Grid Table 3 Accent 4" w:uiPriority="48"/>
            <w:lsdException w:name="Grid Table 4 Accent 4" w:uiPriority="49"/>
            <w:lsdException w:name="Grid Table 5 Dark Accent 4" w:uiPriority="50"/>
            <w:lsdException w:name="Grid Table 6 Colorful Accent 4" w:uiPriority="51"/>
            <w:lsdException w:name="Grid Table 7 Colorful Accent 4" w:uiPriority="52"/>
            <w:lsdException w:name="Grid Table 1 Light Accent 5" w:uiPriority="46"/>
            <w:lsdException w:name="Grid Table 2 Accent 5" w:uiPriority="47"/>
            <w:lsdException w:name="Grid Table 3 Accent 5" w:uiPriority="48"/>
            <w:lsdException w:name="Grid Table 4 Accent 5" w:uiPriority="49"/>
            <w:lsdException w:name="Grid Table 5 Dark Accent 5" w:uiPriority="50"/>
            <w:lsdException w:name="Grid Table 6 Colorful Accent 5" w:uiPriority="51"/>
            <w:lsdException w:name="Grid Table 7 Colorful Accent 5" w:uiPriority="52"/>
            <w:lsdException w:name="Grid Table 1 Light Accent 6" w:uiPriority="46"/>
            <w:lsdException w:name="Grid Table 2 Accent 6" w:uiPriority="47"/>
            <w:lsdException w:name="Grid Table 3 Accent 6" w:uiPriority="48"/>
            <w:lsdException w:name="Grid Table 4 Accent 6" w:uiPriority="49"/>
            <w:lsdException w:name="Grid Table 5 Dark Accent 6" w:uiPriority="50"/>
            <w:lsdException w:name="Grid Table 6 Colorful Accent 6" w:uiPriority="51"/>
            <w:lsdException w:name="Grid Table 7 Colorful Accent 6" w:uiPriority="52"/>
            <w:lsdException w:name="List Table 1 Light" w:uiPriority="46"/>
            <w:lsdException w:name="List Table 2" w:uiPriority="47"/>
            <w:lsdException w:name="List Table 3" w:uiPriority="48"/>
            <w:lsdException w:name="List Table 4" w:uiPriority="49"/>
            <w:lsdException w:name="List Table 5 Dark" w:uiPriority="50"/>
            <w:lsdException w:name="List Table 6 Colorful" w:uiPriority="51"/>
            <w:lsdException w:name="List Table 7 Colorful" w:uiPriority="52"/>
            <w:lsdException w:name="List Table 1 Light Accent 1" w:uiPriority="46"/>
            <w:lsdException w:name="List Table 2 Accent 1" w:uiPriority="47"/>
            <w:lsdException w:name="List Table 3 Accent 1" w:uiPriority="48"/>
            <w:lsdException w:name="List Table 4 Accent 1" w:uiPriority="49"/>
            <w:lsdException w:name="List Table 5 Dark Accent 1" w:uiPriority="50"/>
            <w:lsdException w:name="List Table 6 Colorful Accent 1" w:uiPriority="51"/>
            <w:lsdException w:name="List Table 7 Colorful Accent 1" w:uiPriority="52"/>
            <w:lsdException w:name="List Table 1 Light Accent 2" w:uiPriority="46"/>
            <w:lsdException w:name="List Table 2 Accent 2" w:uiPriority="47"/>
            <w:lsdException w:name="List Table 3 Accent 2" w:uiPriority="48"/>
            <w:lsdException w:name="List Table 4 Accent 2" w:uiPriority="49"/>
            <w:lsdException w:name="List Table 5 Dark Accent 2" w:uiPriority="50"/>
            <w:lsdException w:name="List Table 6 Colorful Accent 2" w:uiPriority="51"/>
            <w:lsdException w:name="List Table 7 Colorful Accent 2" w:uiPriority="52"/>
            <w:lsdException w:name="List Table 1 Light Accent 3" w:uiPriority="46"/>
            <w:lsdException w:name="List Table 2 Accent 3" w:uiPriority="47"/>
            <w:lsdException w:name="List Table 3 Accent 3" w:uiPriority="48"/>
            <w:lsdException w:name="List Table 4 Accent 3" w:uiPriority="49"/>
            <w:lsdException w:name="List Table 5 Dark Accent 3" w:uiPriority="50"/>
            <w:lsdException w:name="List Table 6 Colorful Accent 3" w:uiPriority="51"/>
            <w:lsdException w:name="List Table 7 Colorful Accent 3" w:uiPriority="52"/>
            <w:lsdException w:name="List Table 1 Light Accent 4" w:uiPriority="46"/>
            <w:lsdException w:name="List Table 2 Accent 4" w:uiPriority="47"/>
            <w:lsdException w:name="List Table 3 Accent 4" w:uiPriority="48"/>
            <w:lsdException w:name="List Table 4 Accent 4" w:uiPriority="49"/>
            <w:lsdException w:name="List Table 5 Dark Accent 4" w:uiPriority="50"/>
            <w:lsdException w:name="List Table 6 Colorful Accent 4" w:uiPriority="51"/>
            <w:lsdException w:name="List Table 7 Colorful Accent 4" w:uiPriority="52"/>
            <w:lsdException w:name="List Table 1 Light Accent 5" w:uiPriority="46"/>
            <w:lsdException w:name="List Table 2 Accent 5" w:uiPriority="47"/>
            <w:lsdException w:name="List Table 3 Accent 5" w:uiPriority="48"/>
            <w:lsdException w:name="List Table 4 Accent 5" w:uiPriority="49"/>
            <w:lsdException w:name="List Table 5 Dark Accent 5" w:uiPriority="50"/>
            <w:lsdException w:name="List Table 6 Colorful Accent 5" w:uiPriority="51"/>
            <w:lsdException w:name="List Table 7 Colorful Accent 5" w:uiPriority="52"/>
            <w:lsdException w:name="List Table 1 Light Accent 6" w:uiPriority="46"/>
            <w:lsdException w:name="List Table 2 Accent 6" w:uiPriority="47"/>
            <w:lsdException w:name="List Table 3 Accent 6" w:uiPriority="48"/>
            <w:lsdException w:name="List Table 4 Accent 6" w:uiPriority="49"/>
            <w:lsdException w:name="List Table 5 Dark Accent 6" w:uiPriority="50"/>
            <w:lsdException w:name="List Table 6 Colorful Accent 6" w:uiPriority="51"/>
            <w:lsdException w:name="List Table 7 Colorful Accent 6" w:uiPriority="52"/>
            <w:lsdException w:name="Mention" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Smart Hyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Hashtag" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Unresolved Mention" w:semiHidden="1" w:unhideWhenUsed="1"/>
            <w:lsdException w:name="Smart Link" w:semiHidden="1" w:unhideWhenUsed="1"/>
        </w:latentStyles>
        
        ${Normal}
        ${H1StyleXML}
        ${H2StyleXML}
        ${H3StyleXML}
        ${H4StyleXML}
        ${H5StyleXML}
        ${H6StyleXML}
        ${Heading1Char}
        ${Heading2Char}
        ${Heading3Char}
        ${Heading4Char}
        ${Heading5Char}
        ${Heading6Char}

        ${DefaultParagraphFont}
        ${TableNormal}
        ${NoList}
        ${ListParagraph}

        ${Title}
        ${TitleChar}
        ${Subtitle}
        ${SubtitleChar}

        ${Emphasis}
        ${Strong}
        ${Quote}
        ${QuoteChar}
        ${Hyperlink}
        ${GridTable1Light}

    </w:styles>`

    const theme1XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <a:theme
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office 主题​​">
        <a:themeElements>
            <a:clrScheme name="Office">
                <a:dk1>
                    <a:sysClr val="windowText" lastClr="000000"/>
                </a:dk1>
                <a:lt1>
                    <a:sysClr val="window" lastClr="FFFFFF"/>
                </a:lt1>
                <a:dk2>
                    <a:srgbClr val="44546A"/>
                </a:dk2>
                <a:lt2>
                    <a:srgbClr val="E7E6E6"/>
                </a:lt2>
                <a:accent1>
                    <a:srgbClr val="4472C4"/>
                </a:accent1>
                <a:accent2>
                    <a:srgbClr val="ED7D31"/>
                </a:accent2>
                <a:accent3>
                    <a:srgbClr val="A5A5A5"/>
                </a:accent3>
                <a:accent4>
                    <a:srgbClr val="FFC000"/>
                </a:accent4>
                <a:accent5>
                    <a:srgbClr val="5B9BD5"/>
                </a:accent5>
                <a:accent6>
                    <a:srgbClr val="70AD47"/>
                </a:accent6>
                <a:hlink>
                    <a:srgbClr val="0563C1"/>
                </a:hlink>
                <a:folHlink>
                    <a:srgbClr val="954F72"/>
                </a:folHlink>
            </a:clrScheme>
            <a:fontScheme name="Office">
                <a:majorFont>
                    <a:latin typeface="等线 Light" panose="020F0302020204030204"/>
                    <a:ea typeface=""/>
                    <a:cs typeface=""/>
                    <a:font script="Jpan" typeface="游ゴシック Light"/>
                    <a:font script="Hang" typeface="맑은 고딕"/>
                    <a:font script="Hans" typeface="等线 Light"/>
                    <a:font script="Hant" typeface="新細明體"/>
                    <a:font script="Arab" typeface="Times New Roman"/>
                    <a:font script="Hebr" typeface="Times New Roman"/>
                    <a:font script="Thai" typeface="Angsana New"/>
                    <a:font script="Ethi" typeface="Nyala"/>
                    <a:font script="Beng" typeface="Vrinda"/>
                    <a:font script="Gujr" typeface="Shruti"/>
                    <a:font script="Khmr" typeface="MoolBoran"/>
                    <a:font script="Knda" typeface="Tunga"/>
                    <a:font script="Guru" typeface="Raavi"/>
                    <a:font script="Cans" typeface="Euphemia"/>
                    <a:font script="Cher" typeface="Plantagenet Cherokee"/>
                    <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
                    <a:font script="Tibt" typeface="Microsoft Himalaya"/>
                    <a:font script="Thaa" typeface="MV Boli"/>
                    <a:font script="Deva" typeface="Mangal"/>
                    <a:font script="Telu" typeface="Gautami"/>
                    <a:font script="Taml" typeface="Latha"/>
                    <a:font script="Syrc" typeface="Estrangelo Edessa"/>
                    <a:font script="Orya" typeface="Kalinga"/>
                    <a:font script="Mlym" typeface="Kartika"/>
                    <a:font script="Laoo" typeface="DokChampa"/>
                    <a:font script="Sinh" typeface="Iskoola Pota"/>
                    <a:font script="Mong" typeface="Mongolian Baiti"/>
                    <a:font script="Viet" typeface="Times New Roman"/>
                    <a:font script="Uigh" typeface="Microsoft Uighur"/>
                    <a:font script="Geor" typeface="Sylfaen"/>
                    <a:font script="Armn" typeface="Arial"/>
                    <a:font script="Bugi" typeface="Leelawadee UI"/>
                    <a:font script="Bopo" typeface="Microsoft JhengHei"/>
                    <a:font script="Java" typeface="Javanese Text"/>
                    <a:font script="Lisu" typeface="Segoe UI"/>
                    <a:font script="Mymr" typeface="Myanmar Text"/>
                    <a:font script="Nkoo" typeface="Ebrima"/>
                    <a:font script="Olck" typeface="Nirmala UI"/>
                    <a:font script="Osma" typeface="Ebrima"/>
                    <a:font script="Phag" typeface="Phagspa"/>
                    <a:font script="Syrn" typeface="Estrangelo Edessa"/>
                    <a:font script="Syrj" typeface="Estrangelo Edessa"/>
                    <a:font script="Syre" typeface="Estrangelo Edessa"/>
                    <a:font script="Sora" typeface="Nirmala UI"/>
                    <a:font script="Tale" typeface="Microsoft Tai Le"/>
                    <a:font script="Talu" typeface="Microsoft New Tai Lue"/>
                    <a:font script="Tfng" typeface="Ebrima"/>
                </a:majorFont>
                <a:minorFont>
                    <a:latin typeface="等线" panose="020F0502020204030204"/>
                    <a:ea typeface=""/>
                    <a:cs typeface=""/>
                    <a:font script="Jpan" typeface="游明朝"/>
                    <a:font script="Hang" typeface="맑은 고딕"/>
                    <a:font script="Hans" typeface="等线"/>
                    <a:font script="Hant" typeface="新細明體"/>
                    <a:font script="Arab" typeface="Arial"/>
                    <a:font script="Hebr" typeface="Arial"/>
                    <a:font script="Thai" typeface="Cordia New"/>
                    <a:font script="Ethi" typeface="Nyala"/>
                    <a:font script="Beng" typeface="Vrinda"/>
                    <a:font script="Gujr" typeface="Shruti"/>
                    <a:font script="Khmr" typeface="DaunPenh"/>
                    <a:font script="Knda" typeface="Tunga"/>
                    <a:font script="Guru" typeface="Raavi"/>
                    <a:font script="Cans" typeface="Euphemia"/>
                    <a:font script="Cher" typeface="Plantagenet Cherokee"/>
                    <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
                    <a:font script="Tibt" typeface="Microsoft Himalaya"/>
                    <a:font script="Thaa" typeface="MV Boli"/>
                    <a:font script="Deva" typeface="Mangal"/>
                    <a:font script="Telu" typeface="Gautami"/>
                    <a:font script="Taml" typeface="Latha"/>
                    <a:font script="Syrc" typeface="Estrangelo Edessa"/>
                    <a:font script="Orya" typeface="Kalinga"/>
                    <a:font script="Mlym" typeface="Kartika"/>
                    <a:font script="Laoo" typeface="DokChampa"/>
                    <a:font script="Sinh" typeface="Iskoola Pota"/>
                    <a:font script="Mong" typeface="Mongolian Baiti"/>
                    <a:font script="Viet" typeface="Arial"/>
                    <a:font script="Uigh" typeface="Microsoft Uighur"/>
                    <a:font script="Geor" typeface="Sylfaen"/>
                    <a:font script="Armn" typeface="Arial"/>
                    <a:font script="Bugi" typeface="Leelawadee UI"/>
                    <a:font script="Bopo" typeface="Microsoft JhengHei"/>
                    <a:font script="Java" typeface="Javanese Text"/>
                    <a:font script="Lisu" typeface="Segoe UI"/>
                    <a:font script="Mymr" typeface="Myanmar Text"/>
                    <a:font script="Nkoo" typeface="Ebrima"/>
                    <a:font script="Olck" typeface="Nirmala UI"/>
                    <a:font script="Osma" typeface="Ebrima"/>
                    <a:font script="Phag" typeface="Phagspa"/>
                    <a:font script="Syrn" typeface="Estrangelo Edessa"/>
                    <a:font script="Syrj" typeface="Estrangelo Edessa"/>
                    <a:font script="Syre" typeface="Estrangelo Edessa"/>
                    <a:font script="Sora" typeface="Nirmala UI"/>
                    <a:font script="Tale" typeface="Microsoft Tai Le"/>
                    <a:font script="Talu" typeface="Microsoft New Tai Lue"/>
                    <a:font script="Tfng" typeface="Ebrima"/>
                </a:minorFont>
            </a:fontScheme>
            <a:fmtScheme name="Office">
                <a:fillStyleLst>
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:gradFill rotWithShape="1">
                        <a:gsLst>
                            <a:gs pos="0">
                                <a:schemeClr val="phClr">
                                    <a:lumMod val="110000"/>
                                    <a:satMod val="105000"/>
                                    <a:tint val="67000"/>
                                </a:schemeClr>
                            </a:gs>
                            <a:gs pos="50000">
                                <a:schemeClr val="phClr">
                                    <a:lumMod val="105000"/>
                                    <a:satMod val="103000"/>
                                    <a:tint val="73000"/>
                                </a:schemeClr>
                            </a:gs>
                            <a:gs pos="100000">
                                <a:schemeClr val="phClr">
                                    <a:lumMod val="105000"/>
                                    <a:satMod val="109000"/>
                                    <a:tint val="81000"/>
                                </a:schemeClr>
                            </a:gs>
                        </a:gsLst>
                        <a:lin ang="5400000" scaled="0"/>
                    </a:gradFill>
                    <a:gradFill rotWithShape="1">
                        <a:gsLst>
                            <a:gs pos="0">
                                <a:schemeClr val="phClr">
                                    <a:satMod val="103000"/>
                                    <a:lumMod val="102000"/>
                                    <a:tint val="94000"/>
                                </a:schemeClr>
                            </a:gs>
                            <a:gs pos="50000">
                                <a:schemeClr val="phClr">
                                    <a:satMod val="110000"/>
                                    <a:lumMod val="100000"/>
                                    <a:shade val="100000"/>
                                </a:schemeClr>
                            </a:gs>
                            <a:gs pos="100000">
                                <a:schemeClr val="phClr">
                                    <a:lumMod val="99000"/>
                                    <a:satMod val="120000"/>
                                    <a:shade val="78000"/>
                                </a:schemeClr>
                            </a:gs>
                        </a:gsLst>
                        <a:lin ang="5400000" scaled="0"/>
                    </a:gradFill>
                </a:fillStyleLst>
                <a:lnStyleLst>
                    <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
                        <a:solidFill>
                            <a:schemeClr val="phClr"/>
                        </a:solidFill>
                        <a:prstDash val="solid"/>
                        <a:miter lim="800000"/>
                    </a:ln>
                    <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
                        <a:solidFill>
                            <a:schemeClr val="phClr"/>
                        </a:solidFill>
                        <a:prstDash val="solid"/>
                        <a:miter lim="800000"/>
                    </a:ln>
                    <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
                        <a:solidFill>
                            <a:schemeClr val="phClr"/>
                        </a:solidFill>
                        <a:prstDash val="solid"/>
                        <a:miter lim="800000"/>
                    </a:ln>
                </a:lnStyleLst>
                <a:effectStyleLst>
                    <a:effectStyle>
                        <a:effectLst/>
                    </a:effectStyle>
                    <a:effectStyle>
                        <a:effectLst/>
                    </a:effectStyle>
                    <a:effectStyle>
                        <a:effectLst>
                            <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
                                <a:srgbClr val="000000">
                                    <a:alpha val="63000"/>
                                </a:srgbClr>
                            </a:outerShdw>
                        </a:effectLst>
                    </a:effectStyle>
                </a:effectStyleLst>
                <a:bgFillStyleLst>
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:solidFill>
                        <a:schemeClr val="phClr">
                            <a:tint val="95000"/>
                            <a:satMod val="170000"/>
                        </a:schemeClr>
                    </a:solidFill>
                    <a:gradFill rotWithShape="1">
                        <a:gsLst>
                            <a:gs pos="0">
                                <a:schemeClr val="phClr">
                                    <a:tint val="93000"/>
                                    <a:satMod val="150000"/>
                                    <a:shade val="98000"/>
                                    <a:lumMod val="102000"/>
                                </a:schemeClr>
                            </a:gs>
                            <a:gs pos="50000">
                                <a:schemeClr val="phClr">
                                    <a:tint val="98000"/>
                                    <a:satMod val="130000"/>
                                    <a:shade val="90000"/>
                                    <a:lumMod val="103000"/>
                                </a:schemeClr>
                            </a:gs>
                            <a:gs pos="100000">
                                <a:schemeClr val="phClr">
                                    <a:shade val="63000"/>
                                    <a:satMod val="120000"/>
                                </a:schemeClr>
                            </a:gs>
                        </a:gsLst>
                        <a:lin ang="5400000" scaled="0"/>
                    </a:gradFill>
                </a:bgFillStyleLst>
            </a:fmtScheme>
        </a:themeElements>
        <a:objectDefaults/>
        <a:extraClrSchemeLst/>
        <a:extLst>
            <a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">
                <thm15:themeFamily
                    xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/>
                </a:ext>
            </a:extLst>
        </a:theme>`

    const webSettingsXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:webSettings
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
        xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
        xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
        xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
        xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
        xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
        xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh">
        <w:optimizeForBrowser/>
        <w:allowPNG/>
    </w:webSettings>`

    const word = zip.folder('word');
    word.folder('_rels').file('document.xml.rels', wordRelsXML);
    word.file('document.xml', documentXML);
    word.file('fontTable.xml', fontTableXML);
    word.file('numbering.xml', numberingXML);
    word.file('settings.xml', settingsXML);
    word.file("styles.xml", stylesXML);
    word.folder('theme').file('theme1.xml', theme1XML)
    word.file("webSettings.xml", webSettingsXML);
}

function escapeXml(unsafe) {
    return unsafe.replace(/[<>&"']/g, function (c) {
      switch (c) {
        case '<': return '&lt;';
        case '>': return '&gt;';
        case '&': return '&amp;';
        case '"': return '&quot;';
        case "'": return '&apos;';
      }
    });
  }
  