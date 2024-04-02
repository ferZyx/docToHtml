"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var pandoc_ts_1 = require("pandoc-ts");
var outs = [
    { name: "html", format: "html", fname: "sample.html" }
];
var pandocInstance = new pandoc_ts_1.Pandoc("docx", outs);
pandocInstance.convertFile("original.docx", function (result, err) {
    if (err) {
        console.error(err);
    }
    if (result) {
        console.log(result.html);
    }
});
