import ADMZip from "adm-zip";
import xml2js from "xml2js";

type Props = {
  URL?: string;
  Local?: string;
  Replacements: Replacements;
  Out?: string;
};

export type CellData = { Address: { Row: number; Col: number }; Name: string; Value: string };

export type ExcelData = {
  Name: string;
  Workbook: string;
  Data: CellData[];
};

export type Table = {
  Name: string;
  Headers: string[];
  Data: string[][];
};

export type Replacements = {
  [key: string]: string | ExcelData[] | Table[] | undefined;
  XLSX?: ExcelData[];
  TABLES?: Table[];
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

const ReplaceBasic = async (Entries: ADMZip.IZipEntry[], Replacements: Replacements, Category: string) => {
  const Slides = Entries.filter((E) => E.entryName.startsWith(`ppt/${Category}/`) && !E.entryName.endsWith(".rels"));
  for (let x = 0; x < Slides.length; x++) {
    const TargetSlide = Slides[x];
    let NewContent = TargetSlide.getData().toString();
    let Changed = false;
    for (const [Key, Value] of Object.entries(Replacements)) {
      Changed = true;
      if (NewContent.includes(Key)) {
        if (NewContent.includes(Tag(Key))) {
          if (typeof Value === "string") {
            const StringValue = Value as string;
            NewContent = NewContent.replaceAll(Tag(Key), EncodeXml(StringValue));
          }
        } else {
          let SlideContent = await xml2js.parseStringPromise(NewContent);
          for (let [i, C] of SlideContent["p:sld"]["p:cSld"].entries()) {
            const SPs = C["p:spTree"][0]["p:sp"];
            for (let [j, SP] of SPs.entries()) {
              const Ps = SP["p:txBody"][0]["a:p"];
              for (let [k, P] of Ps.entries())
                if (P) {
                  if (P["a:r"]?.length > 2) {
                    const T = P["a:r"].map((R: any) => R["a:t"][0]);
                    if (T.join("") === Tag(Key)) {
                      SlideContent["p:sld"]["p:cSld"][i]["p:spTree"][0]["p:sp"][j]["p:txBody"][0]["a:p"][k]["a:r"][0][
                        "a:t"
                      ][0] = Value;
                      delete SlideContent["p:sld"]["p:cSld"][i]["p:spTree"][0]["p:sp"][j]["p:txBody"][0]["a:p"][k][
                        "a:r"
                      ][1];
                      delete SlideContent["p:sld"]["p:cSld"][i]["p:spTree"][0]["p:sp"][j]["p:txBody"][0]["a:p"][k][
                        "a:r"
                      ][2];
                    }
                  }
                }
            }

            const builder = new xml2js.Builder();
            NewContent = builder.buildObject(SlideContent);
          }
        }
      }
    }
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
    x < ChartArea?.["c:ser"]?.[0]?.["c:val"]?.[0]?.["c:numRef"]?.[0]?.["c:numCache"]?.[0]?.["c:pt"]?.length || 0;
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
const RandomNumber = () => {
  let str = Math.floor(100000000 + Math.random() * 900000000).toString();
  while (str.length < 9) {
    str = "0" + str;
  }
  return str;
};
const GridColBasic = (Width: number, uri?: string) => {
  const Out: any = {
    $: { w: Width.toString() },
  };
  Out["a:extLst"] = [
    {
      "a:ext": [
        {
          $: {
            uri: uri,
          },
          "a16:colId": [
            {
              $: { "xmlns:a16": "http://schemas.microsoft.com/office/drawing/2014/main", val: RandomNumber() },
            },
          ],
        },
      ],
    },
  ];
  return Out;
};
const TableColBasic = {
  "a:txBody": [{ "a:bodyPr": [{}], "a:lstStyle": [{}], "a:p": [{ "a:r": [{ "a:t": [""] }] }] }],
};
const TableRowBasic = (Height: number, uri: string, Cols: number) => {
  return {
    $: { h: Height.toString() },
    "a:tc": Array(Cols).fill(TableColBasic),
    "a:extLst": [
      {
        "a:ext": [
          {
            $: { uri: uri },
            // "a16:rowId": {
            //   $: { "xmlns:a16": "http://schemas.microsoft.com/office/drawing/2014/main", val: RandomNumber() },
            // },
          },
        ],
      },
    ],
  };
};
const ReplaceTables = async (Entries: ADMZip.IZipEntry[], Replacements: Table[]) => {
  const Slides = Entries.filter((E) => E.entryName.startsWith(`ppt/slides/`) && !E.entryName.endsWith(".rels"));
  for (let iSlide = 0; iSlide < Slides.length; iSlide++) {
    const TargetSlide = Slides[iSlide];
    let NewContent = TargetSlide.getData().toString();
    let Changed = false;
    for (const [Key, TableDetails] of Object.entries(Replacements)) {
      Changed = true;
      if (NewContent.includes(Key)) {
        let SlideContent = await xml2js.parseStringPromise(NewContent);
        for (let [iSlideContent, C] of SlideContent["p:sld"]["p:cSld"].entries()) {
          const GFs = C["p:spTree"][0]["p:graphicFrame"];
          if (!GFs) break;
          for (let [iGF, GF] of GFs.entries()) {
            for (let [iG, G] of GF["a:graphic"].entries()) {
              for (let [iGD, GD] of G["a:graphicData"].entries()) {
                if (!GD["a:tbl"]) break;
                for (let [iT, T] of GD["a:tbl"].entries()) {
                  const TableContent = T["a:tr"]
                    .map((tr: any) =>
                      tr["a:tc"]
                        .map((tc: any) =>
                          tc["a:txBody"][0]["a:p"][0]["a:r"]
                            ? tc["a:txBody"][0]["a:p"][0]["a:r"].map((r: any) => (r["a:t"] ? r["a:t"][0] : "")).join("")
                            : ""
                        )
                        .join("")
                    )
                    .join("");
                  if (TableContent.includes(Tag(TableDetails.Name))) {
                    let TargetTable =
                      SlideContent["p:sld"]["p:cSld"][iSlideContent]["p:spTree"][0]["p:graphicFrame"][iGF]["a:graphic"][
                        iG
                      ]["a:graphicData"][iGD]["a:tbl"][iT];
                    let Original = {
                      ...SlideContent["p:sld"]["p:cSld"][iSlideContent]["p:spTree"][0]["p:graphicFrame"][iGF][
                        "a:graphic"
                      ][iG]["a:graphicData"][iGD]["a:tbl"][iT],
                    };
                    const GridColumnDefinition = { ...TargetTable["a:tblGrid"][0]["a:gridCol"][0] };
                    const TableRowHDef = TargetTable["a:tr"][0];
                    const ColDiff = TableDetails.Headers.length - TargetTable["a:tr"][1]["a:tc"].length;
                    for (let c = 0; c < ColDiff; c++) {
                      TargetTable["a:tblGrid"][0]["a:gridCol"].push({
                        $: { w: GridColumnDefinition["$"].w },
                        "a:extLst": [
                          {
                            "a:ext": [
                              {
                                $: {
                                  uri: GridColumnDefinition["a:extLst"][0]["a:ext"][0]["$"].uri,
                                },
                                "a16:colId": [
                                  {
                                    $: {
                                      "xmlns:a16": "http://schemas.microsoft.com/office/drawing/2014/main",
                                      val: GridColumnDefinition["a:extLst"][0]["a:ext"][0]["a16:colId"][0]["$"].val,
                                    },
                                  },
                                ],
                              },
                            ],
                          },
                        ],
                      });
                      TargetTable["a:tr"][0]["a:tc"].push({
                        "a:txBody": [{ "a:bodyPr": [""], "a:lstStyle": [""], "a:p": [{ "a:r": [{ "a:t": [""] }] }] }],
                        "a:tcPr": [
                          {
                            "a:solidFill": [
                              {
                                "a:srgbClr": [
                                  {
                                    $: {
                                      val: TableRowHDef["a:tc"][0]["a:tcPr"][0]["a:solidFill"][0]["a:srgbClr"][0]["$"]
                                        .val,
                                    },
                                  },
                                ],
                              },
                            ],
                          },
                        ],
                      });
                      TargetTable["a:tr"][1]["a:tc"].push({
                        "a:txBody": [{ "a:bodyPr": [""], "a:lstStyle": [""], "a:p": [{ "a:r": [{ "a:t": [""] }] }] }],
                      });
                      TargetTable["a:tr"][2]["a:tc"].push({
                        "a:txBody": [{ "a:bodyPr": [""], "a:lstStyle": [""], "a:p": [{ "a:r": [{ "a:t": [""] }] }] }],
                      });
                    }
                    for (let c = 0; c < TableDetails.Headers.length; c++) {
                      TargetTable["a:tr"][0]["a:tc"][c]["a:txBody"][0]["a:p"][0]["a:r"] = [
                        { "a:t": [TableDetails.Headers[c]] },
                      ];
                    }
                    const TableRowGDefs = [{ ...TargetTable["a:tr"][1] }, { ...TargetTable["a:tr"][2] }];
                    for (let r = 0; r < TableDetails.Data.length; r++) {
                      for (let c = 0; c < TableDetails.Data[r].length; c++) {
                        if (r + 1 >= TargetTable["a:tr"].length) {
                          TargetTable["a:tr"].push({
                            $: { h: TableRowGDefs[r % 2]["$"].h },
                            "a:tc": [],
                            "a:extLst": [
                              {
                                "a:ext": [
                                  {
                                    $: { uri: TableRowGDefs[r % 2]["a:extLst"][0]["a:ext"][0]["$"].uri },
                                    "a16:rowId": {
                                      $: {
                                        "xmlns:a16": "http://schemas.microsoft.com/office/drawing/2014/main",
                                        val: RandomNumber(),
                                      },
                                    },
                                  },
                                ],
                              },
                            ],
                          });
                        }
                        if (!TargetTable["a:tr"][r + 1]["a:tc"][c])
                          TargetTable["a:tr"][r + 1]["a:tc"].push({
                            "a:txBody": [
                              {
                                "a:bodyPr": [""],
                                "a:lstStyle": [""],
                                "a:p": [
                                  {
                                    "a:pPr": TableRowGDefs[r % 2]["a:tc"][c]["a:txBody"][0]["a:p"][0]["a:pPr"],
                                    "a:r": [{ "a:t": [""] }],
                                  },
                                ],
                              },
                            ],
                          });
                        if (r + 1 < TargetTable["a:tr"].length) {
                          TargetTable["a:tr"][r + 1]["a:tc"][c]["a:txBody"][0]["a:p"][0]["a:r"] = [
                            { "a:t": [TableDetails.Data[r][c]] },
                          ];
                          delete TargetTable["a:tr"][r + 1]["a:tc"][c]["a:txBody"][0]["a:p"][0]["a:endParaRPr"];
                        }
                      }
                    }
                  }
                }
              }
            }
          }
          const builder = new xml2js.Builder();
          NewContent = builder.buildObject(SlideContent);
        }
      }
    }
    if (Changed) {
      const EntryNames = Entries.map((E) => E.entryName);
      const TargetIndex = EntryNames.indexOf(TargetSlide.entryName);
      Entries[TargetIndex].setData(Buffer.from(NewContent));
    }
  }
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
  if (!!Replacements.TABLES) await ReplaceTables(Entries, Replacements.TABLES);
  await ReplaceBasic(Entries, BasicReplacements, "slides");
  await ReplaceBasic(Entries, BasicReplacements, "charts");
  if (Replacements.XLSX) {
    await ReplaceXLSX(Entries, Replacements.XLSX ?? []);
    await ReplaceCharts(Entries, Replacements.XLSX ?? []);
  }
  if (Out) return PPTX.writeZip(Out);
  const OutBuffer = PPTX.toBuffer();
  return OutBuffer;
};

export default msotr;
