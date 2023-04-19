import './style.css'
import javascriptLogo from './javascript.svg'
import viteLogo from '/vite.svg'
import { setupCounter } from './counter.js'

import { saveDocx } from "./src/index";
import { GenerateDocx } from "./src/docx-xml";
import { docx } from "./src/docx-os";

const htmlString = '<div class="markdown prose w-full break-words dark:prose-invert light"><p>Sure, here are some of the most commonly used Markdown tags:</p><h1>Headers</h1><h1>H1</h1><h2>H2</h2><h3>H3</h3><h4>H4</h4><h5>H5</h5><h6>H6</h6><h1>Emphasis</h1><p><em>italic</em>\n<strong>bold</strong>\n<em><strong>bold and italic</strong></em></p><h1>Lists</h1><h2>Unordered list</h2><ul><li>Item 1</li><li>Item 2</li><li>Item 3</li></ul><h2>Ordered list</h2><ol><li>Item 1</li><li>Item 2</li><li>Item 3</li></ol><h1>Links</h1><p><a href="URL" target="_new">Link text</a></p><h1>Images</h1><p>![Alt text](image URL)</p><h1>Code</h1><p><code>Inline code</code></p><pre><div class="bg-black rounded-md mb-4"><div class="flex items-center relative text-gray-200 bg-gray-800 px-4 py-2 text-xs font-sans justify-between rounded-t-md"><span>css</span><button class="flex ml-auto gap-2"><svg stroke="currentColor" fill="none" stroke-width="2" viewBox="0 0 24 24" stroke-linecap="round" stroke-linejoin="round" class="h-4 w-4" height="1em" width="1em" xmlns="http://www.w3.org/2000/svg"><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"></path><rect x="8" y="2" width="8" height="4" rx="1" ry="1"></rect></svg>Copy code</button></div><div class="p-4 overflow-y-auto"><code class="!whitespace-pre hljs language-css"><span class="hljs-selector-tag">Code</span> block\n</code></div></div></pre><h1>Horizontal line</h1><hr><h1>Blockquotes</h1><blockquote><p>Quote text</p></blockquote><h1>Tables</h1><table><thead><tr><th>Column 1</th><th>Column 2</th><th>Column 3</th></tr></thead><tbody><tr><td>Row 1, Column 1</td><td>Row 1, Column 2</td><td>Row 1, Column 3</td></tr><tr><td>Row 2, Column 1</td><td>Row 2, Column 2</td><td>Row 2, Column 3</td></tr></tbody></table></div>';

const htmlString2 = document.querySelector('#export-container')

document.querySelector('#btn-save').addEventListener('click',()=>{
  saveDocx(htmlString)
})

document.querySelector('#btn-save-gpt').addEventListener('click',()=>{
  const g = new GenerateDocx(htmlString2.innerHTML)
  g.save()
})
document.querySelector('#btn-save-gpt-docx').addEventListener('click',()=>{
  docx({DOM:htmlString2})
})
