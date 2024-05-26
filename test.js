import {docToDocxByLibreOffice, wordToHtmlByPandoc} from "./wordToHtmlConverter.js";
import fs from 'fs';

fs.mkdirSync('tests/html-libre', {recursive: true});
fs.mkdirSync('tests/html-pandoc', {recursive: true});
fs.mkdirSync('tests/docx', {recursive: true});

(async () => {

    await docToDocxByLibreOffice('tests/word/' + 'main.doc', './tests/docx');

    await wordToHtmlByPandoc('tests/docx/main.docx', 'tests/html-pandoc/' + 'main.html')

})()