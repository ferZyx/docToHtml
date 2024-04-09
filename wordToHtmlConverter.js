import { spawn } from "child_process";
export async function wordToHtmlByLibreOffice(wordName, outdir, args = []) {
    return new Promise((resolve, reject) => {
        const commandPrompt = ['--headless', '--convert-to', 'html:HTML:EmbedImages', wordName, '--outdir', outdir, ...args];
        let libreoffice = spawn("libreoffice", commandPrompt);
        libreoffice.stdout.on("data", (data) => {
            console.log('stdout:', data.toString());
        });
        libreoffice.on("error", (err) => {
            reject(err);
        });
        libreoffice.on("exit", (code, signal) => {
            if (code !== 0) {
                reject(new Error('Ошибка конвертации файла. Код: ' + code + ' ' + signal));
            }
            else {
                resolve();
            }
        });
    });
}
