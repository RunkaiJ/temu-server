const xlsx = require("xlsx");

// ------------------ Hardcoded Defaults ------------------
const HARDCODE = {
    EntryType: "01",
    ForwardingID: "2568210",
    ConsigneeId: "2567704",
    RelatedParties: "Y",
    "Mode of Transport": "40",
    SellingName: "August Amber HK Limited",
    SellingStreetAddress:
        "Unit 417, 4/F Lippo Ctr Tower Two, No.89 Queensway, Admiralty, Hong Kong",
    SellingCity: "HongKong",
    SellingPostalCode: "",
    SellingCountry: "HK",
    Unit: "PCS",
    "Preparer Port": "4701",
};

// ------------------ Location Mapping ------------------
const LOCATION_MAP = {
    4701: "EAT5",
    2720: "WBH9",
    2801: "W0B3",
    3901: "HBT1",
    5501: "SE04",
    5206: "LEG0",
    1704: "L543",
    "0417": "AAN5",
    3029: "WBU6",
};

// ------------------ Manufacturer Code Mapping ------------------
const RAW_MANUF_CODES = [
    ["GUANGZHOUPEISHANFUZHUANGGONGYINGLIANGUANICO.LTD", "CNPEISHA2263GUA"],
    ["GanZhouPanHong Technology Co., LTD", "CNGANTEC1026GAN"],
    ["Wenzhoushensengongju Co., Ltd.", "CNSHESENWEN"],
    ["Wenzhouyihewanjuzhizao Co., Ltd.", "CNWENCO497WEN"],
    ["Jinhua Jindong District Liju E-commerce Firm", "CNJINJIN5547JIN"],
    ["GUANGXIPINGNANXIANYUHANKEJI CO..LTD", "CNGUACOL2202GUI"],
    ["SHENZHEN BEAUTIFUL STORY TRADING CO., LTD.", "CNSHEBEA3241SHE"],
    ["BANANA TOOTH CLOTHING CO.LTD", "CNBANTOO463GUA"],
    ["SHENZHENJIECHANGSHENGINDUSTRIALCO..LTD", "CNSHE683SHE"],
    // ["SHENZHENTUOWEIDIANZISHANGWUCO.LTD.", "CNSHE201SHE"],
    ["Wenzhouheshengsujiao Co., Ltd.", "CNHESSUJWEN"],
    // ["Wenzhouchangchuangshangmao Co., Ltd.", "CNWENZHOWEN"],
    ["FOSHANBEISHIYUFUSHI CO., LTD.", "CNFOSCO6120FOS"],
    ["SHANDONGSIBANGGONGJU CO.LTD", "CNSHACOL9115LIN"],
    ["WENZHOUFEIMUCLOTHINGCO.LTD", "CNWEN937WEN"],
    ["ZHENGZHOUAMIERMAOYI CO.LTD", "CNZHECOL157ZHE"],
    ["PUJIANGXIANFULIUMANMAOYICO.LTD.", "CNPUJ772JIN"],
    // ["Wenzhoushimikaixieye Co., Ltd.", "CNWENZHOWEN"],
    ["rizhaozhulongfangzhiyouxiangongsi", "CNRIZ624RIZ"],
    ["YIWU XINMAN JEWELRY CO. LTD", "CNYIWXIN7062JIN"],
    // ["GanZhouPanHongTechnology", "CNBANTOO463GUA"],
    ["LINZHOUJINPENGSHANGMAO CO..LTD.", "CNLINCOL1126ANY"],
    ["NO.1177,BINHAI, LONGWAN, WENZHOU.", "CNNO1BIN303JIN"],
    ["XINYUSHIXIANNVHUQUQIANHEBAIHUIZHIXIECHANG", "CNXIN9104XIN"],
    // ["Wenzhoushifengyiwujinzhipin Co., Ltd.", "CNWENZHOWEN"],
    ["XINGCHENGHUIMEICLOTHINGMANUFACTURINGFACTORY", "CNXIN2100HUL"],
    ["FOSHANSHIMMUPINGFUSHI O..LTD.", "CNFOSOLT1090FOS"],
    ["Haomiao.electronic.commerce co.,ltd", "CNHAOCOL912TAI"],
    ["ZHONGQINMAOYIGUANGZHOU CO..LTD", "CNZHOCOL493GUA"],
    ["GUANGZHOUSHILINGLINGQIKEJIFAZHANCO.LTD.", "CNGUA2330GUA"],
    ["NO.5, PETROLEUM ROAD", "CNNO5PET5XIA"],
    ["PUJIANG MEMORY TOY CO., LTD", "CNPUJMEM8282JIN"],
    ["JINHUASHIQINRONGMAOYI CO.LTD", "CNJINCOL360JIN"],
    ["ANHUI HEYE NETWORK TECHNOLOGY CO., LTD", "CNANHHEY2105HEF"],
    ["SHENZHENJIAMIWANGLUOISHUCO..LTD", "CNSHE179SHE"],
    ["SHENZHENSHIANRUNTAIKEJICO.LTD.", "CNSHE1353SHE"],
    ["GUANGZHOUTONGRUNTRADINGCO..LTD", "CNGUA162JIE"],
    ["GUANGZHOU AICHUANGJINCHUKOUMAOYI CO., LTD.", "CNGUAAIC3183GUA"],
    ["Fujian Chengrui Shoes Industry Co. , Ltd.", "CNFUJCHE814FUZ"],
    ["PANAN JINGYA HORTICULTURAL SUPPLIES FACTORY", "CNPANJIN202JIN"],
    ["HUAIYUANXINHAODIANZISHANGWU CO.LTD.", "CNHUACOL104BEN"],
    ["GUANGZHOUTONGRUNTRADINGCO..LTD", "CNGUA162JIE"],
    ["HUIYU WOMEN'S SHOE FACTORY", "CNHUIWOM2282XIN"],
    ["Foshan Shanmu clothing Co., LTD", "CNFOSSHA2450FOS"],
    ["HUIZHOU DUDU PET PRODUCTS CO., ITD", "CNHUIDUD234HUI"],
    [
        "Hengyang City Coarse Tea and Rice Catering Culture Ltd",
        "CNHENCIT153HEN",
    ],
    ["JINHUAGUOYOUHOMEFURNISHINGCO.LTD", "CNJIN890JIN"],
    ["TAIZHOUSHIDAHONGBAOGONGYIPIN CO.LTD.", "CNTAICOL413TAI"],
    ["Foshanshixuntuodianzico.Ltd", "CNFOS2119FOS"],
    ["YIWUSHIYOUKEJINCHUKOUCO.LTD.", "CNYIW2310JIN"],
    ["GUANGZHAUYUNNIFUSHICO.LTD.", "CNGUA1076GUA"],
    ["ZHEJIANG MENGDA IMPORT AND EXPORT CO., LTD.", "CNZHEMEN1188JIN"],
    ["SHENZHENSHIDONGHEXINKEJI.CO,.LTD", "CNSHELTD238SHE"],
    [
        "MOUDING COUNTY KUNI DEPARTMENT STORE SOLE PROPRIETORSHIP",
        "CNMOUCOU2744CHU",
    ],
    ["SHANDONGTAHAIWANGLUOKEJI CO.LTD.", "CNSHACOL1595RIZ"],
    ["GUANGZHOUHUIWANGDIANZISHANGWUCO..LTD", "CNGUA850DON"],
    ["Xingcheng Kaisheng Garment Co., Ltd.", "CNKAIGARXIN"],
    ["Xingcheng Kaiwei Garment Co., Ltd.", "CNXINKAI6543HUL"],
    ["Xingcheng Kairuida Garment Co", "CNKAIGARHUL"],

];


const keyName = (s = "") =>
    s
        .toString()
        .replace(/[^a-z0-9]/gi, "")
        .toLowerCase();

const MANUF_CODE_MAP = RAW_MANUF_CODES.reduce((acc, [name, code]) => {
    acc[keyName(name)] = code;
    return acc;
}, {});

function getManufacturerCode(name) {
    return MANUF_CODE_MAP[keyName(name)] || "";
}

// Header order copied from your template (must match EXACTLY):
const HEADERS = [
    "EntryNo",
    "EntryType",
    "GroupIdentifier",
    "ForwardingID",
    "ConsigneeId",
    "RelatedParties",
    "DescOfMerchandise",
    "HTS",
    "HTSValue",
    "UnladingPort",
    "EntryDate",
    "ImportDate",
    "Mode of Transport",
    "ManufacturerCode",
    "ManufacturerName",
    "ManufacturerStreetAddress",
    "ManufacturerCity",
    "ManufacturerCountry",
    "ManufacturerPostalCode",
    "ManufacturerProvince",
    "SellingMID",
    "SellingName",
    "SellingStreetAddress",
    "SellingCity",
    "SellingState",
    "SellingPostalCode",
    "SellingCountry",
    "InvoiceNumber",
    "HTSQty",
    "HTSQty2",
    "HTS-1",
    "HTS-2",
    "HTS-3",
    "Airline 3 digit code",
    "Master Bill Number",
    "House AWB",
    "Manifest Qty Piece count",
    "Unit",
    "Description",
    "Date of Export",
    "Country Of Origin",
    "Country of Export",
    "Arrival Date",
    "Location of Goods",
    "Carrier Code",
    "Voyage Flight No",
    "Arrival Airport",
    "Preparer Port",
    "Remote Port",
    "STATE OF DESTINATION",
    "SteelCountryOfMeltAndPour",
    "SteelCountryOfMeltAndPourAppCode",
    "PrimaryAluminumCountryOfSmelt",
    "PrimaryAluminumCountryOfSmeltAppCode",
    "SecondaryAluminumCountryOfSmelt",
    "SecondaryAluminumCountryOfSmeltAppCode",
    "AluminumCountryOfCastCode",
    "FDAPRODUCTCODE",
    "FDAPROGRAMCODE",
    "FDAPROCESSINGCODE",
    "FDAINTENDEDUSECODE",
    "FDABRANDNAME",
    "FDAARRIVALTIME",
    "FDANAME",
    "FDAADDRESS",
    "FDACITY",
    "FDACOUNTRY",
    "FDAREGISTRATIONNUMBERTYPE",
    "FDAREGISTRATIONNUMBER",
    "VESSELNAMEORRAILCARNO",
    "COMPLIANCECODE1",
    "COMPLIANCECODE1VALUE",
    "COMPLIANCECODE2",
    "COMPLIANCECODE2VALUE",
    "COMPLIANCECODE3",
    "COMPLIANCECODE3VALUE",
    "COMPLIANCECODE4",
    "COMPLIANCECODE4VALUE",
    "LACEYDECLARATIONSIGNEDDATE",
    "LACEYLINEVALUE",
    "LACEYENTITYROLECODE",
    "LACEYENTITYNAME",
    "LACEYENTITYEMAIL",
    "LACEYENTITYPHONE",
    "LACEYENTITYNAME-1",
    "LACEYENTITYEMAIL-1",
    "LACEYENTITYPHONE-1",
    "LACEYACTIVEINGREDIENT",
    "LACEYNAMEOFELEMENT",
    "LACEYQUANTITYOFELEMENT",
    "LACEYUNITOFMEASURE",
    "LACEYPERCENTOFELEMENT",
    "LACEYGENUSNAME",
    "LACEYSPECIESNAME",
    "LACEYSUBSPECIESNAME",
    "LACEYSPECIESCODE",
    "LACEYDESCRIPTIONCODE",
    "LACEYSOURCETYPECODE",
    "LACEYCOUNTRYCODE",
    "LACEYPRODUCTCOMPONENT-1",
    "LACEYACTIVEINGREDIENT-1",
    "LACEYNAMEOFELEMENT-1",
    "LACEYQUANTITYOFELEMENT-1",
    "LACEYUNITOFMEASURE-1",
    "LACEYPERCENTOFELEMENT-1",
    "LACEYGENUSNAME-1",
    "LACEYSPECIESNAME-1",
    "LACEYSUBSPECIESNAME-1",
    "LACEYSPECIESCODE-1",
    "LACEYDESCRIPTIONCODE-1",
    "LACEYSOURCETYPECODE-1",
    "LACEYCOUNTRYCODE-1",
    "LACEYGEOGRAPHICLOCATION",
    "LACEYPROCESSINGSTARTDATE",
    "LACEYPROCESSINGTYPECODE",
    "LACEYPROCESSINGDESCRIPTION",
    "LACEYCONTAINERNUMBER",
    "LACEYCONTAINERNUMBER-1",
    "LACEYLICENSETYPE",
    "LACEYTRANSACTIONTYPE",
    "LACEYLICENSENUMBER",
    "LACEYLPCODATETYPE",
    "LACEYLPCODATE",
    "LACEYLICENSETYPE-1",
    "LACEYTRANSACTIONTYPE-1",
    "LACEYLICENSENUMBER-1",
    "LACEYLPCODATETYPE-1",
    "LACEYLPCODATE-1",
];

// ------------------ Helpers ------------------
function toMMDDYYYY(iso) {
    if (!iso) return "";
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(iso)) return iso;
    const [y, m, d] = iso.split("-");
    return `${m.padStart(2, "0")}/${d.padStart(2, "0")}/${y}`;
}

function normalizeHeader(str) {
    return str
        ?.toString()
        .replace(/\u00A0/g, " ")
        .trim()
        .toLowerCase();
}

// ------------------ Main Conversion ------------------
async function convertTemuExcel(manifestBuffer, formData) {
    // Read manifest file
    const wb = xlsx.read(manifestBuffer, { type: "buffer" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const raw = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    // Find header row by "tracking_number"
    const headerRowIndex = raw.findIndex((row) =>
        row.some(
            (c) => c && c.toString().toLowerCase().trim() === "tracking_number"
        )
    );
    if (headerRowIndex < 0) {
        throw new Error("Header row with 'tracking_number' not found");
    }

    const headers = raw[headerRowIndex];
    const dataRows = raw
        .slice(headerRowIndex + 1)
        .filter((r) => r[headers.indexOf("description")]);

    // Column index helpers
    const idx = (name) =>
        headers.findIndex((h) => normalizeHeader(h) === name.toLowerCase());

    const colTracking = idx("tracking_number");
    const colGroup = idx("group_no");
    const colDesc = idx("description");
    const colHTS = idx("harmonization_code");
    const colValue = idx("total_value");
    const colMfgName = idx("manufacture_name");
    const colMfgAddr = idx("manufacture_address");
    const colMfgCity = idx("manufacture_city");
    const colMfgZip = idx("manufacture_zip_code");
    const colMfgCountry = idx("manufacture_country");
    const colQty = idx("quantity");
    const colCountryOrigin = idx("country_of_origin");

    // pull PGA/FDA mapping columns from source
    const colPGAProductCode = idx("PGA_Product_Code");
    const colProcessingCode = idx("processing_code");
    const colProgramCode = idx("program_code");

    // Prepare user inputs
    const portCode = (formData.portCode || "").trim();
    const locationOfGoods = LOCATION_MAP[portCode] || "";
    const formattedDate = toMMDDYYYY(formData.date);

    // Build output rows
    const out = [];

    for (let i = 0; i < dataRows.length; i++) {
        const r = dataRows[i];

        const originFlag = r[colCountryOrigin];
        const trackingNumber = r[colTracking] || `at row ${i + 1}`;

        if (!originFlag) {
            throw new Error(
                `Missing country_of_origin for tracking number: ${trackingNumber}`
            );
        }

        const newRow = Object.fromEntries(HEADERS.map((h) => [h, ""]));

        // Hardcoded fields
        Object.assign(newRow, HARDCODE);

        // Manifest-driven fields
        newRow.GroupIdentifier = r[colGroup];
        newRow.InvoiceNumber = r[colTracking];

        const desc = r[colDesc];
        newRow.DescOfMerchandise = desc;
        newRow.Description = desc;

        newRow.HTS = r[colHTS];
        newRow.HTSValue = parseFloat(r[colValue]) || 0;
        newRow.HTSQty = r[colQty];
        newRow["Manifest Qty Piece count"] = r[colQty];

        const origin = r[colCountryOrigin];
        newRow["Country Of Origin"] = origin;
        newRow["Country of Export"] = origin;

        // map PGA columns to FDA columns
        const pgaProduct =
            (colPGAProductCode >= 0 ? r[colPGAProductCode] : "")
                ?.toString()
                .trim() || "";
        const processingCode =
            (colProcessingCode >= 0 ? r[colProcessingCode] : "")
                ?.toString()
                .trim() || "";
        const programCode =
            (colProgramCode >= 0 ? r[colProgramCode] : "")?.toString().trim() ||
            "";

        newRow.FDAPRODUCTCODE = pgaProduct;
        newRow.FDAPROCESSINGCODE = processingCode;
        newRow.FDAPROGRAMCODE = programCode;

        //  set intended use when all three exist
        if (pgaProduct && processingCode && programCode) {
            newRow.FDAINTENDEDUSECODE = 130.029;
        }

        // Manufacturer logic
        const rawName = r[colMfgName];
        const code = getManufacturerCode(rawName);
        if (code) {
            newRow.ManufacturerCode = code;
            newRow.ManufacturerName = "";
            newRow.ManufacturerStreetAddress = "";
            newRow.ManufacturerCity = "";
            newRow.ManufacturerCountry = "";
            newRow.ManufacturerPostalCode = "";
        } else {
            newRow.ManufacturerName = rawName;
            newRow.ManufacturerStreetAddress = r[colMfgAddr];
            newRow.ManufacturerCity = r[colMfgCity];
            newRow.ManufacturerCountry = r[colMfgCountry];
            newRow.ManufacturerPostalCode = r[colMfgZip];
        }
        newRow.ManufacturerProvince = "";

        // User-provided and derived fields
        newRow.UnladingPort = portCode;
        newRow["Arrival Airport"] = portCode;
        newRow["Remote Port"] = portCode;
        newRow["Location of Goods"] = locationOfGoods;
        newRow["STATE OF DESTINATION"] = formData.destinationState || "";
        newRow["Carrier Code"] = formData.carrierCode || "";
        newRow["Voyage Flight No"] = formData.voyageFlightNo || "";
        newRow["House AWB"] = formData.houseAWB || "";

        newRow.EntryDate = formattedDate;
        newRow.ImportDate = formattedDate;
        newRow["Arrival Date"] = formattedDate;
        newRow["Date of Export"] = formattedDate;

        newRow["Airline 3 digit code"] = formData.airlineCode || "";
        newRow["Master Bill Number"] = formData.masterBillNumber || "";

        out.push(newRow);
    }

    // Generate workbook
    const sheetOut = xlsx.utils.json_to_sheet(out, { header: HEADERS });
    const wbOut = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wbOut, sheetOut, "Sheet1");
    // const buffer = xlsx.write(wbOut, { type: "buffer", bookType: "xlsx" });
    const buffer = xlsx.write(wbOut, {
        type: "buffer",
        bookType: "xlsx",
        bookSST: true, // build a shared string table (smaller files)
        compression: true, // zip compression on the XLSX
    });

    return { buffer };
}

module.exports = { convertTemuExcel };
