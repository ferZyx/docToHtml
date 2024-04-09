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
        await wordToHtmlByLibreOffice('word/' + wordName ,'./html-libre');
    }

    if (runtime === "pandocone") {
        const wordName = process.argv[3];
        const htmlName = process.argv[4];
        await wordToHtmlByPandoc('tests/word/' + wordName, 'tests/html-pandoc/' + htmlName);
    }

    if (runtime === "libremany") {
        await Promise.allSettled([
            wordToHtmlByLibreOffice('word/1.docx', './html-libre'),
            wordToHtmlByLibreOffice('word/2.docx', './html-libre'),
            wordToHtmlByLibreOffice('word/3.docx', './html-libre'),
            wordToHtmlByLibreOffice('word/4.docx', './html-libre'),
            wordToHtmlByLibreOffice('word/5.docx', './html-libre'),
        ])
    }

})
().then(() => {
    const endTime = new Date().getTime();
    console.log(`Time spent: ${endTime - startTime} ms. `);
})
