import {wordToHtmlByLibreOffice} from "../wordToHtmlConverter.js";

const startTime = new Date().getTime();

await (async () => {
    const runtime = process.argv[2]
    console.log(runtime)
    if (runtime === "libreOne") {
        const wordName = process.argv[3];
        try {
            await wordToHtmlByLibreOffice(wordName, './html-libre');
            console.log('Conversion completed successfully');
        } catch (err) {
            console.error('Error at test.js:', err);
        }
    }

    if (runtime === "libreMany") {
        try {
            wordToHtmlByLibreOffice("1.doc", './html-libre').then(() => {
                    console.log('Conversion completed successfully')
                }
            )
            wordToHtmlByLibreOffice("2.docx", './html-libre').then(() => {
                    console.log('Conversion completed successfully')
                }
            )
            wordToHtmlByLibreOffice("3.docx", './html-libre').then(() => {
                    console.log('Conversion completed successfully')
                }
            )
        } catch (err) {
            console.error('Error at test.js:', err);
        }
    }

})();
const endTime = new Date().getTime();
console.log(`Time spent: ${endTime - startTime} ms. `);
