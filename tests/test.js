import {wordToHtmlByLibreOffice} from "../wordToHtmlConverter.js";

(async () => {
    const startTime = new Date().getTime();

    const runtime = process.argv[2]
    if (runtime === "libreOne"){
        const wordName = process.argv[3];
        try {
            await wordToHtmlByLibreOffice(wordName, './html-libre');
            console.log('Conversion completed successfully');
        } catch (err) {
            console.error('Error at test.js:', err);
        }
    }

    if (runtime === "libreMany"){
        try {
            wordToHtmlByLibreOffice("1.doc", './html-libre');
            wordToHtmlByLibreOffice("2.docx", './html-libre');
            wordToHtmlByLibreOffice("3.docx", './html-libre');
        } catch (err) {
            console.error('Error at test.js:', err);
        }
    }

    const endTime = new Date().getTime();
    console.log(`Time spent: ${endTime - startTime} ms. Runtime: ${runtime}`);
})();