
/**
 * http://officeopenxml.com/WPsampleDoc.php
 */

// 导入jszip库
import JSZip from 'jszip';


// 生成docx文件的函数
export function generateDocx(html) {

  // 解析 HTML 字符串为 DOM 对象
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");

  // 创建一个新的JSZip实例
  const zip = new JSZip();

  let xml = ``;

  for (const node of doc.body.childNodes) {
    xml += traverse(node);
  }
  xml = `<w:p><w:r><w:t>Sure, here are some of the most commonly used Markdown tags:</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1" /></w:pPr><w:r><w:t>Headers</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1" /></w:pPr><w:r><w:t>H1</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading2" /></w:pPr><w:r><w:t>H2</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading3" /></w:pPr><w:r><w:t>H3</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading4" /></w:pPr><w:r><w:t>H4</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading5" /></w:pPr><w:r><w:t>H5</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading6" /></w:pPr><w:r><w:t>H6</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1" /></w:pPr><w:r><w:t>Emphasis</w:t></w:r></w:p><w:p><w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t><w:rPr><w:b/></w:rPr><w:t>bold</w:t><w:rPr><w:i/></w:rPr><w:rPr><w:b/></w:rPr><w:t>bold and italic</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1" /></w:pPr><w:r><w:t>Lists</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading2" /></w:pPr><w:r><w:t>Unordered list</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading2" /></w:pPr><w:r><w:t>Ordered list</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1" /></w:pPr><w:r><w:t>Links</w:t></w:r></w:p><w:p><w:r></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1" /></w:pPr><w:r><w:t>Images</w:t></w:r></w:p><w:p><w:r><w:t>![Alt text](image URL)</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1" /></w:pPr><w:r><w:t>Code</w:t></w:r></w:p><w:p><w:r></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1" /></w:pPr><w:r><w:t>Horizontal line</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1" /></w:pPr><w:r><w:t>Blockquotes</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1" /></w:pPr><w:r><w:t>Tables</w:t></w:r></w:p>`
  console.info('xml-=',xml)
  // return

  // for (const node of doc.body.childNodes) {
  //   // 处理文本节点
  //   if (node.nodeType === Node.TEXT_NODE) {
  //     xml += `<w:p><w:r><w:t>${node.textContent}</w:t></w:r></w:p>`;
  //   }

  //   // 处理加粗标签
  //   else if (node.nodeType === Node.ELEMENT_NODE && (node.nodeName === "STRONG" || node.nodeName === "B")) {
  //     xml += `<w:p><w:r>`;
  //     for (const child of node.childNodes) {
  //       if (child.nodeType === Node.TEXT_NODE) {
  //         xml += `<w:rPr><w:b/></w:rPr><w:t>${child.textContent}</w:t>`;
  //       }
  //     }
  //     xml += `</w:r></w:p>`;
  //   }

  //   // 处理斜体标签
  //   else if (node.nodeType === Node.ELEMENT_NODE && (node.nodeName === "EM" || node.nodeName === "I")) {
  //     xml += `<w:p><w:r>`;
  //     for (const child of node.childNodes) {
  //       if (child.nodeType === Node.TEXT_NODE) {
  //         xml += `<w:rPr><w:i/></w:rPr><w:t>${child.textContent}</w:t>`;
  //       }
  //     }
  //     xml += `</w:r></w:p>`;
  //   }

  //   // 处理段落标签
  //   else if (node.nodeType === Node.ELEMENT_NODE && node.nodeName === "P") {
  //     xml += `<w:p><w:r>`;
  //     for (const child of node.childNodes) {
  //       if (child.nodeType === Node.TEXT_NODE) {
  //         xml += `<w:t>${child.textContent}</w:t>`;
  //       } else if (
  //         child.nodeType === Node.ELEMENT_NODE &&
  //         child.nodeName === "STRONG"
  //       ) {
  //         for (const grandchild of child.childNodes) {
  //           if (grandchild.nodeType === Node.TEXT_NODE) {
  //             xml += `<w:rPr><w:b/></w:rPr><w:t>${grandchild.textContent}</w:t>`;
  //           }
  //         }
  //       }
  //       // 处理其他标签
  //       // ...
  //     }
  //     xml += `</w:r></w:p>`;
  //   }


  //   // 处理 H1-H6 标签
  //   else if (node.nodeType === Node.ELEMENT_NODE && /^H[1-6]$/.test(node.nodeName)) {
  //     const headingLevel = node.nodeName.substring(1);
  //     xml += `<w:p><w:pPr><w:pStyle w:val="Heading${headingLevel}" /></w:pPr><w:r><w:t>`;
  //     for (const child of node.childNodes) {
  //       if (child.nodeType === Node.TEXT_NODE) {
  //         xml += `${child.textContent}`;
  //       }
  //     }
  //     xml += `</w:t></w:r></w:p>`;
  //   }



  //   // 处理其他标签
  //   // ...

  // }

  // 将构建的xml字符串添加到JSZip实例中
  setContentTypes(zip)
  setRels(zip)
  setDocProps(zip)
  setWord(zip, xml)


  // 将生成的docx文件下载到本地
  zip.generateAsync({ type: 'blob' }).then(function (content) {
    const a = document.createElement('a');
    const url = URL.createObjectURL(content);
    a.href = url;
    a.download = 'docx-template.docx';
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    }, 0);
  });
}

function traverse(node) {
  // 处理文本节点
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent.trim() === '' ? '' : `<w:t>${node.textContent}</w:t>`;
  }

  // 处理加粗标签
  else if (
    node.nodeType === Node.ELEMENT_NODE &&
    (node.nodeName === "STRONG" || node.nodeName === "B")
  ) {
    let childrenXml = "";
    for (const child of node.childNodes) {
      childrenXml += traverse(child);
    }
    return `<w:rPr><w:b/></w:rPr>${childrenXml}`;
  }
  // 处理斜体标签
  else if (
    node.nodeType === Node.ELEMENT_NODE &&
    (node.nodeName === "EM" || node.nodeName === "I")
  ) {
    let childrenXml = "";
    for (const child of node.childNodes) {
      childrenXml += traverse(child);
    }
    return `<w:rPr><w:i/></w:rPr>${childrenXml}`;
  }

  // 处理段落标签
  else if (node.nodeType === Node.ELEMENT_NODE && node.nodeName === "P") {
    let childrenXml = "";
    for (const child of node.childNodes) {
      childrenXml += traverse(child);
    }
    return `<w:p><w:r>${childrenXml}</w:r></w:p>`;
  }

  // 处理 H1-H6 标签
  else if (node.nodeType === Node.ELEMENT_NODE && /^H[1-6]$/.test(node.nodeName)) {
    const headingLevel = node.nodeName.substring(1);
    let childrenXml = "";
    for (const child of node.childNodes) {
      childrenXml += traverse(child);
    }
    return `<w:p><w:pPr><w:pStyle w:val="Heading${headingLevel}" /></w:pPr><w:r>${childrenXml}</w:r></w:p>`;
  }

  // 处理其他标签
  // ...

  return "";
}

function setContentTypes(zip) {
  // 构建docx所需的xml字符串
  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
   <Types
       xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
       <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
       <Default Extension="xml" ContentType="application/xml"/>
       <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
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
      <TotalTime>0</TotalTime>
      <Pages>1</Pages>
      <Words>0</Words>
      <Characters>${charCount}</Characters>
      <Application>Microsoft Office Word</Application>
      <DocSecurity>0</DocSecurity>
      <Lines>1</Lines>
      <Paragraphs>${pCount}</Paragraphs>
      <ScaleCrop>false</ScaleCrop>
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

function setWord(zip, xml) {
  const wordRelsXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships
        xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
        <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
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
        <w:font w:name="等线">
            <w:altName w:val="DengXian"/>
            <w:panose1 w:val="02010600030101010101"/>
            <w:charset w:val="86"/>
            <w:family w:val="auto"/>
            <w:pitch w:val="variable"/>
            <w:sig w:usb0="A00002BF" w:usb1="38CF7CFA" w:usb2="00000016" w:usb3="00000000" w:csb0="0004000F" w:csb1="00000000"/>
        </w:font>
        <w:font w:name="Times New Roman">
            <w:panose1 w:val="02020603050405020304"/>
            <w:charset w:val="00"/>
            <w:family w:val="roman"/>
            <w:pitch w:val="variable"/>
            <w:sig w:usb0="E0002EFF" w:usb1="C000785B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
        </w:font>
        <w:font w:name="等线 Light">
            <w:panose1 w:val="02010600030101010101"/>
            <w:charset w:val="86"/>
            <w:family w:val="auto"/>
            <w:pitch w:val="variable"/>
            <w:sig w:usb0="A00002BF" w:usb1="38CF7CFA" w:usb2="00000016" w:usb3="00000000" w:csb0="0004000F" w:csb1="00000000"/>
        </w:font>
    </w:fonts>`

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
        <w:proofState w:grammar="clean"/>
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
            <w:rsidRoot w:val="00DE143E"/>
            <w:rsid w:val="009F7532"/>
            <w:rsid w:val="00DE143E"/>
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
        <w14:docId w14:val="1C23AE55"/>
        <w15:chartTrackingRefBased/>
        <w15:docId w15:val="{EB32034E-C7EB-45E9-87EE-E1406F519F7D}"/>
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

        ${Title}
        ${TitleChar}
        ${Subtitle}
        ${SubtitleChar}

        ${Emphasis}
        ${Strong}
        ${Quote}
        ${QuoteChar}

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
  word.file('settings.xml', settingsXML);
  word.file("styles.xml", stylesXML);
  word.folder('theme').file('theme1.xml', theme1XML)
  word.file("webSettings.xml", webSettingsXML);
}