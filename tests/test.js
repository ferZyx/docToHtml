import {wordToHtmlByLibreOffice} from "../wordToHtmlConverter.js";

const startTime = new Date().getTime();

(async () => {
    const runtime = process.argv[2]
    console.log(runtime)

    if (runtime === "libreone") {
        const wordName = process.argv[3];
        await wordToHtmlByLibreOffice(wordName, './html-libre');
    }

    if (runtime === "libremany") {
        await Promise.all([
            wordToHtmlByLibreOffice('word/1.doc', './html-libre'),
            wordToHtmlByLibreOffice('word/2.docx', './html-libre'),
            wordToHtmlByLibreOffice('word/3.docx', './html-libre'),
        ])
    }

})
().then(() => {
    const endTime = new Date().getTime();
    console.log(`Time spent: ${endTime - startTime} ms. `);
})
