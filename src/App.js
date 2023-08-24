import React from "react";
import "./styles.css";
import Docxtemplater from "docxtemplater";
import ImageModule from "../image-module/index.js";
import PizZip from "pizzip";
import PizZipUtils from "pizzip/utils/index.js";
import { saveAs } from "file-saver";
import file from "../assets/PredefinedTemplates.pptx";
import testData from "./testData";

function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}

export default function App() {
  const renderDoc = () => {
    loadFile(file, (error: any, content: any) => {
      if (error) {
        console.error(error);
        return;
      }

      const opts = {};
      opts.centered = false;
      opts.getImage = function (tagValue: any, tagName: any) {
        return new Promise((resolve, reject) => {
          PizZipUtils.getBinaryContent(tagValue, (error, content) => {
            if (error) {
              return reject(error);
            }
            return resolve(content);
          });
        });
      };
      opts.getSize = function (img: any, tagValue: any, tagName: any) {
        // FOR FIXED SIZE IMAGE :
        return [400, 400];
      };

      const imageModule = new ImageModule(opts);

      const zip = new PizZip(content);
      const doc = new Docxtemplater()
        .loadZip(zip)
        .attachModule(imageModule)
        .compile();

      doc.resolveData(testData).then(() => {
        doc.render();

        const out = doc.getZip().generate({
          type: "blob",
          mimeType:
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        });

        saveAs(out, "test");
      });
    });
  };

  return (
    <div className="App">
      <h1>Generate PowerPoint Template</h1>
      <button onClick={renderDoc}>Generate Template</button>
    </div>
  );
}
