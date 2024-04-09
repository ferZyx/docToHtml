import {wordToHtmlByLibreOffice, wordToHtmlByPandoc} from "./wordToHtmlConverter.js";
import fs from 'fs';

const startTime = new Date().getTime();

fs.mkdirSync('tests/html-libre', { recursive: true });
fs.mkdirSync('tests/html-pandoc', { recursive: true });

(async () => {
    const runtime = process.argv[2]
    console.log(runtime)

    if (runtime === "libreone") {
        const wordName = process.argv[3];
        await wordToHtmlByLibreOffice('tests/word/' + wordName ,'./tests/html-libre');
    }

    if (runtime === "pandocone") {
        const wordName = process.argv[3];
        const htmlName = process.argv[4];
        await wordToHtmlByPandoc('tests/word/' + wordName, 'tests/html-pandoc/' + htmlName);
    }

    if (runtime === "libremany") {
        await Promise.allSettled([
            wordToHtmlByLibreOffice('tests/word/1.docx', './tests/html-libre'),
            wordToHtmlByLibreOffice('tests/word/2.docx', './tests/html-libre'),
            wordToHtmlByLibreOffice('tests/word/3.docx', './tests/html-libre'),
            wordToHtmlByLibreOffice('tests/word/4.docx', './tests/html-libre'),
            wordToHtmlByLibreOffice('tests/word/5.docx', './tests/html-libre'),
        ])
    }

})
().then(() => {
    const endTime = new Date().getTime();
    console.log(`Time spent: ${endTime - startTime} ms. `);
})
