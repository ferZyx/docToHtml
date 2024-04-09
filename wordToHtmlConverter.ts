import {ChildProcessWithoutNullStreams, spawn} from "child_process";

export async function wordToHtmlByLibreOffice(wordName: string, outdir: string, args: string[] = []): Promise<void> {
    return new Promise<void>((resolve, reject): void => {
        const commandPrompt: string[] = ['--headless', '--convert-to', 'html:HTML:EmbedImages', wordName, '--outdir', outdir, ...args];
        let libreoffice: ChildProcessWithoutNullStreams = spawn("libreoffice", commandPrompt);

        libreoffice.stdout.on("data", (data: Buffer) => {
            console.log('stdout:', data.toString());
        });

        libreoffice.on("error", (err) => {
            reject(err);
        });

        libreoffice.on("exit", (code, signal) => {
            if (code !== 0) {
                reject(new Error('Ошибка конвертации файла. Код: ' + code + ' ' + signal));
            } else {
                resolve();
            }
        });
    });
}