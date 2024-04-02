import { Pandoc, PandocOutFormat } from "pandoc-ts";

const outs: PandocOutFormat = { name: "html", format:"html", fname: "sample.html" };

const pandocInstance = new Pandoc("docx", outs);

// pandocInstance.convertFile("input.docx", (result, err) => {
//     if (err) {
//         console.error(err);
//     }
//     if (result) {
//         console.log(result.html);
//     }
// });

(async () => {
    await pandocInstance.convertFileAsync("input.docx");
})