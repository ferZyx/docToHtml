import {ChildProcessWithoutNullStreams, spawn} from "child_process";

async function convertWordToHtmlAsync(wordName: string, args: string[] = []): Promise<void> {
    return new Promise<void>((resolve, reject): void => {
        const commandPrompt: string[] = ['--headless', '--convert-to', 'html:HTML:EmbedImages', wordName, '--outdir', 'html', ...args];
        let libreoffice: ChildProcessWithoutNullStreams = spawn("libreoffice", commandPrompt);

        libreoffice.stdout.on("data", (data: Buffer) => {
            console.log('stdout:', data.toString());
        });

        libreoffice.on("error", (err) => {
            console.error('Error:', err);
            reject(new Error('Error converting file'));
        });

        libreoffice.on("exit", (code, signal) => {
            if (code !== 0) {
                console.error('Ошибка конвертации:', code, signal);
                reject(new Error('Ошибка конвертации файла'));
            } else {
                console.log('Конвертация завершена успешно');
                resolve();
            }
        });
    });
}

convertWordToHtmlAsync('original.docx')
    .then(() => {
        console.log('Conversion DOCX! completed successfully');
        convertWordToHtmlAsync('original.doc')
            .then(() => {
                console.log('Conversion DOC! completed successfully');
            })
            .catch((error) => {
                console.error('Conversion failed:', error);
            });

    })
    .catch((error) => {
        console.error('Conversion failed:', error);
    });

