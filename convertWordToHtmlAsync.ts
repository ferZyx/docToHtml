import {ChildProcessWithoutNullStreams, spawn} from "child_process";

export async function convertWordToHtmlAsync(wordName: string, outdir: string, args: string[] = []): Promise<void> {
    return new Promise<void>((resolve, reject): void => {
        const commandPrompt: string[] = ['--headless', '--convert-to', 'html:HTML:EmbedImages', wordName, '--outdir', outdir, ...args];
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

convertWordToHtmlAsync('1.doc', 'html').then(() => {
        console.log('Конвертация завершена успешно 1')
    }
).catch((err) => {
    console.error('Ошибка конвертации 1:', err);
});

convertWordToHtmlAsync('2.docx', 'html').then(() => {
        console.log('Конвертация завершена успешно 2')
    }
).catch((err) => {
    console.error('Ошибка конвертации 2:', err);
});

convertWordToHtmlAsync('3.docx', 'html').then(() => {
        console.log('Конвертация завершена успешно 3')
    }
).catch((err) => {
    console.error('Ошибка конвертации: 3', err);
});
