import ADMZip from "adm-zip";
import xml2js from "xml2js";
import fs from "fs";

type CellData = { Address: { Row: number; Col: number }; Name: string; Value: string };
type ExcelData = {
  Name: string;
  Workbook: string;
  Data: CellData[];
};
type Tables = {
  Name: string;
  Headers: string[];
  Row: {
    Col: string[];
  };
};
type Replacements = {
  [key: string]: string | ExcelData[] | Tables[] | undefined;
  XLSX?: ExcelData[];
  TABLES?: Tables[];
};

const TagOpen = "{";
const TagClose = "}";
const Tag = (Text: string) => `${TagOpen}${Text}${TagClose}`;

const EncodeXml = (Unsafe: string) => {
  return Unsafe.replaceAll(/&/g, "&amp;")
    .replaceAll(/</g, "&lt;")
    .replaceAll(/>/g, "&gt;")
    .replaceAll(/"/g, "&quot;")
    .replaceAll(/'/g, "&apos;");
};

const DecodeXml = (Unsafe: string) => {
  return Unsafe.replaceAll(/&/g, "&amp;")
    .replaceAll(/&lt;/g, "<")
    .replaceAll(/&gt;/g, ">")
    .replaceAll(/&quot;/g, '"')
    .replaceAll(/&apos;/g, "'");
};

const ReplaceBasic = async (Entries: ADMZip.IZipEntry[], Replacements: Replacements) => {
  const Slides = Entries.filter((E) => E.entryName.startsWith("ppt/slides/") && !E.entryName.endsWith(".rels"));
  for (let x = 0; x < Slides.length; x++) {
    const TargetSlide = Slides[x];
    let NewContent = TargetSlide.getData().toString();
    let SlideContent = await xml2js.parseStringPromise(NewContent);
    let Changed = false;
    Object.entries(Replacements).map(([Key, Value]) => {
      Changed = true;
      if (NewContent.includes(Key)) {
        if (NewContent.includes(Tag(Key))) {
          if (typeof Value === "string") {
            const StringValue = Value as string;
            NewContent = NewContent.replaceAll(Tag(Key), EncodeXml(StringValue));
          }
        } else {
          for (let [i, C] of SlideContent["p:sld"]["p:cSld"].entries()) {
            const Ps = C["p:spTree"][0]["p:sp"][0]["p:txBody"][0]["a:p"];
            for (let [j, P] of Ps.entries())
              if (P) {
                if (P["a:r"].length > 2) {
                  const T = P["a:r"].map((R: any) => R["a:t"][0]);
                  if (T.join("") === Tag(Key)) {
                    SlideContent["p:sld"]["p:cSld"][i]["p:spTree"][0]["p:sp"][0]["p:txBody"][0]["a:p"][j]["a:r"][0][
                      "a:t"
                    ][0] = Value;
                    delete SlideContent["p:sld"]["p:cSld"][i]["p:spTree"][0]["p:sp"][0]["p:txBody"][0]["a:p"][j][
                      "a:r"
                    ][1];
                    delete SlideContent["p:sld"]["p:cSld"][i]["p:spTree"][0]["p:sp"][0]["p:txBody"][0]["a:p"][j][
                      "a:r"
                    ][2];
                  }
                }
              }

            const builder = new xml2js.Builder();
            NewContent = builder.buildObject(SlideContent);
          }
        }
      }
    });
    if (Changed) {
      const EntryNames = Entries.map((E) => E.entryName);
      const TargetIndex = EntryNames.indexOf(TargetSlide.entryName);
      Entries[TargetIndex].setData(Buffer.from(NewContent));
    }
  }
};

const ReplaceXLSXData = async (SheetData: string, Data: CellData[]) => {
  let Worksheet = await xml2js.parseStringPromise(SheetData);
  for (let D of Data) {
    if (!!Worksheet.worksheet.sheetData[0].row[D.Address.Row]?.c?.[D.Address.Col].v) {
      Worksheet.worksheet.sheetData[0].row[D.Address.Row].c[D.Address.Col].v = [D.Value];
    }
  }
  const builder = new xml2js.Builder();
  return builder.buildObject(Worksheet);
};

const ReplaceXLSX = async (Entries: ADMZip.IZipEntry[], Replacements: ExcelData[]) => {
  if (Replacements.length === 0) return;
  const SheetLocation = "xl/worksheets/sheet1.xml";
  const WBs = Entries.filter((E) => E.entryName.startsWith("ppt/embeddings/"));
  const WBNames = Replacements.map((WB) => WB.Workbook);
  for (let x = 0; x < WBs.length; x++) {
    const TargetSheetIndex = WBNames.indexOf(WBs[x].name) ?? -1;
    if (TargetSheetIndex > -1) {
      const TargetSheet = WBs[x];
      const XLSX = new ADMZip(TargetSheet.getData());
      const XLSXEntries = XLSX.getEntries();
      const SheetData = XLSXEntries.filter((X) => X.entryName === SheetLocation)[0]
        .getData()
        .toString();
      const NewContent = await ReplaceXLSXData(SheetData, Replacements[TargetSheetIndex].Data);
      const XLSXEntryNames = XLSXEntries.map((E) => E.entryName);
      const TargetXLSXIndex = XLSXEntryNames.indexOf(SheetLocation);
      XLSXEntries[TargetXLSXIndex].setData(NewContent);
      const EntryNames = Entries.map((E) => E.entryName);
      const TargetIndex = EntryNames.indexOf(TargetSheet.entryName);
      Entries[TargetIndex].setData(XLSX.toBuffer());
    }
  }
};

const ReplaceChartData = (ChartDetails: any, NewChartData: ExcelData) => {
  const PlotArea = ChartDetails["c:chartSpace"]["c:chart"][0]["c:plotArea"][0];
  let ChartArea;
  let ChartType = "";
  if (PlotArea["c:pieChart"]) {
    ChartArea = PlotArea["c:pieChart"][0];
    ChartType = "c:pieChart";
  }
  if (PlotArea["c:barChart"]) {
    ChartArea = PlotArea["c:barChart"][0];
    ChartType = "c:barChart";
  }
  if (!NewChartData) return;
  for (
    let x = 0;
    x < ChartArea?.["c:ser"]?.[0]?.["c:val"]?.[0]?.["c:numRef"]?.[0]?.["c:numCache"]?.[0]?.["c:pt"].length || 0;
    x++
  )
    if (
      !!NewChartData.Data[x] &&
      !!ChartArea["c:ser"][0]["c:val"][0]["c:numRef"][0]["c:numCache"][0]["c:pt"][x]["c:v"][0]
    )
      ChartArea["c:ser"][0]["c:val"][0]["c:numRef"][0]["c:numCache"][0]["c:pt"][x]["c:v"][0] =
        NewChartData.Data[x].Value;
  if (ChartDetails?.["c:chartSpace"]?.["c:chart"]?.[0]?.["c:plotArea"]?.[0]?.[ChartType]?.[0])
    ChartDetails["c:chartSpace"]["c:chart"][0]["c:plotArea"][0][ChartType][0] = ChartArea;
  return ChartDetails;
};

const ReplaceCharts = async (Entries: ADMZip.IZipEntry[], Replacements: ExcelData[]) => {
  const Charts = Entries.filter((E) => E.entryName.startsWith("ppt/charts/chart"));
  for (let x = 0; x < Charts.length; x++) {
    const TargetChart = Charts[x];
    const Chart = TargetChart.getData();
    if (!Replacements?.[x]) return;
    const ChartDetails = ReplaceChartData(await xml2js.parseStringPromise(Chart.toString()), Replacements?.[x]);
    const builder = new xml2js.Builder();
    const EntryNames = Entries.map((E) => E.entryName);
    const TargetIndex = EntryNames.indexOf(TargetChart.entryName);
    Entries[TargetIndex].setData(builder.buildObject(ChartDetails));
  }
  return;
};

type Props = {
  URL?: string;
  Local?: string;
  Replacements: Replacements;
  Out?: string;
};

const msotr = async ({ URL, Local, Replacements, Out }: Props) => {
  if (!URL && !Local) return;
  let PPTX: ADMZip | undefined = undefined;
  if (!!URL) {
    const Res = await fetch(URL);
    if (!Res.ok) return;
    PPTX = new ADMZip(Buffer.from(await Res.arrayBuffer()));
  }
  if (!!Local) {
    PPTX = new ADMZip(Local);
  }
  if (!PPTX) return;
  const Entries = PPTX.getEntries();
  let BasicReplacements = { ...Replacements };
  delete BasicReplacements.XLSX;
  delete BasicReplacements.TABLES;
  await ReplaceBasic(Entries, BasicReplacements);
  if (Replacements.XLSX) {
    await ReplaceXLSX(Entries, Replacements.XLSX ?? []);
    await ReplaceCharts(Entries, Replacements.XLSX ?? []);
  }
  if (Out) return PPTX.writeZip(Out);
  const OutBuffer = PPTX.toBuffer();
  return OutBuffer;
};

export default msotr;
for (let i = 0; i < process.argv.length; ++i) {
  console.log(`index ${i} argument -> ${process.argv[i]}`);
}
if (!!process.argv[3]) {
  const f = fs.readFileSync(process.argv[2]);
  const Replacements = JSON.parse(f.toString());
  if (process.argv[3].startsWith("http"))
    msotr({ URL: process.argv[3], Replacements: Replacements, Out: process.argv[4] });
  else msotr({ Local: process.argv[3], Replacements: Replacements, Out: process.argv[4] });
} else {
  console.log("USEAGE: npm run start [ReplacementJSON] [Template.pptx] [Output.pptx]");
}
