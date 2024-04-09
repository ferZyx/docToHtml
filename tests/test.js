import {wordToHtmlByLibreOffice} from "../wordToHtmlConverter.js";

(async () => {
    const startTime = new Date().getTime();

    const runtime = process.argv[2]
    if (runtime === "libre"){
        const wordName = process.argv[3];
        try {
            await wordToHtmlByLibreOffice(wordName, './tests/html');
            console.log('Conversion completed successfully');
        } catch (err) {
            console.error('Error at test.js:', err);
        }
    }

    const endTime = new Date().getTime();
    console.log(`Time spent: ${endTime - startTime} ms. Runtime: ${runtime}`);
})();