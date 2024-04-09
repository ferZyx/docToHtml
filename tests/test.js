import {wordToHtmlByLibreOffice} from "../wordToHtmlConverter.js";

const startTime = new Date().getTime();

(async () => {
    const runtime = process.argv[2]
    console.log(runtime)
    if (runtime === "libreone") {
        const wordName = process.argv[3];
        try {
            await wordToHtmlByLibreOffice(wordName, './html-libre');
            console.log('Conversion completed successfully');
        } catch (err) {
            console.error('Error at test.js:', err);
        }
    }

    if (runtime === "libremany") {
        try {
            await Promise.all([
                wordToHtmlByLibreOffice('word/1.doc', './html-libre'),
                wordToHtmlByLibreOffice('word/2.docx', './html-libre'),
                wordToHtmlByLibreOffice('word/3.docx', './html-libre'),
            ])

        } catch (err) {
            console.error('Error at test.js:', err);
        }
    }

})().then(() => {
    const endTime = new Date().getTime();
    console.log(`Time spent: ${endTime - startTime} ms. `);
})
