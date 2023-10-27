import fs from "fs";
import msotr from "./index";

// for (let i = 0; i < process.argv.length; ++i) {
//   console.log(`index ${i} argument -> ${process.argv[i]}`);
// }
if (!!process.argv[3]) {
  const f = fs.readFileSync(process.argv[2]);
  const Replacements = JSON.parse(f.toString());
  if (process.argv[3].startsWith("http"))
    msotr({ URL: process.argv[3], Replacements: Replacements, Out: process.argv[4] });
  else msotr({ Local: process.argv[3], Replacements: Replacements, Out: process.argv[4] });
} else {
  console.log("USEAGE: npm run start [ReplacementJSON] [Template.pptx] [Output.pptx]");
}
