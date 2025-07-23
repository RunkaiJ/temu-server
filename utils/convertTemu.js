const xlsx = require("xlsx");

// ------------------ helpers ------------------
// function formatDateMMDDYYYY(dateString) {
//   const d = new Date(dateString);
//   if (isNaN(d)) return "";
//   const mm = String(d.getMonth() + 1).padStart(2, "0");
//   const dd = String(d.getDate()).padStart(2, "0");
//   const yyyy = d.getFullYear();
//   return `${mm}/${dd}/${yyyy}`;
// }

function toMMDDYYYY(iso) {
  if (!iso) return "";
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(iso)) return iso; // already formatted
  const [y, m, d] = iso.split("-");
  return `${m.padStart(2,"0")}/${d.padStart(2,"0")}/${y}`;
}

function findCellAfterLabel(rows, label) {
    const needle = label.toLowerCase();
    for (const row of rows) {
        const idx = row.findIndex(
            (c) => c && c.toString().toLowerCase().includes(needle)
        );
        if (idx !== -1 && row[idx + 1]) {
            return row[idx + 1].toString().trim();
        }
    }
    return "";
}

function parseFromBL(blStr) {
    // "272-74824400"  -> ["272", "74824400"]
    const m = blStr.match(/(\d{3})-([0-9]+)/);
    return m ? { airlineCode: m[1], masterBillNumber: m[2] } : {};
}

function parseFromPO(poStr) {
    // "_272-74824400_..." -> ["272","74824400"]
    const m = poStr.match(/_(\d+)-(\d+)_/);
    return m ? { airlineCode: m[1], masterBillNumber: m[2] } : {};
}

// map Arrival Airport (unlading port code) -> Location of Goods
const LOCATION_MAP = {
  "4701": "EAT5", // JFK
  "2720": "WBH9", // LAX
  "2801": "W0B3", // SFO
  "3901": "HBT1", // ORD
  "5501": "SE04", // DFW
  "5206": "LEG0", // MIA
  "1704": "L543", // ATL
  "0417": "AAN5", // BOS
  "3029": "WBU6", // SEA
};

// These are the ONLY fields you said must be hardcoded
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
  SellingPostalCode: "999077",
  SellingCountry: "HK",
  Unit: "PCS",
  "Preparer Port": "4701", // if this is different, change here
};

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

function normalize(str) {
  return str?.toString().replace(/\u00A0/g, " ").trim().toLowerCase();
}

async function convertTemuExcel(combineBuffer, formData) {
    const combineWb = xlsx.read(combineBuffer, { type: "buffer" });
    const combineSheet = combineWb.Sheets[combineWb.SheetNames[0]];
    const combineRaw = xlsx.utils.sheet_to_json(combineSheet, {
        header: 1,
        defval: "",
    });

    // Combine header row (index 9) & data rows
    const headers = combineRaw[9];
    const itemRows = combineRaw.slice(10).filter((r) => r[2] && r[3]);

    const getColumnIndex = (name) =>
        headers.findIndex((h) => normalize(h) === normalize(name));

    const colHTS = getColumnIndex("HTS Code");
    const colQTY = getColumnIndex("QTY");
    const colSubtotal = getColumnIndex("Subtotal (USD)");
    const colManufName = getColumnIndex("Manufacturer Name");
    const colManufAddr = getColumnIndex("Manufacturer Address");
    const colManufCity = getColumnIndex("Manufacturer City");
    const colManufCountry = getColumnIndex("Manufacturer Country");
    const colManufPostal = getColumnIndex("Manufacturer Address Postal Code");
    const colCountryOrigin = getColumnIndex("Country of Origin");
    const colCommodity = getColumnIndex("commodity");

    // Invoice number (search top area)
    let invoiceNumber = "";
    for (const row of combineRaw) {
        const idx = row.findIndex(
            (c) => c && c.toString().toLowerCase().includes("invoice number")
        );
        if (idx !== -1 && row[idx + 1]) {
            invoiceNumber = row[idx + 1].toString().trim();
            break;
        }
    }

    // Get airlineCode and masterBillNumber from "PO No" line
    let airlineCode = "";
    let masterBillNumber = "";
    const blVal = findCellAfterLabel(combineRaw, "b/l no");
    if (blVal) {
        const parsed = parseFromBL(blVal);
        airlineCode = parsed.airlineCode || "";
        masterBillNumber = parsed.masterBillNumber || "";
    }
    
    //   const formattedDate = formatDateMMDDYYYY(formData.date);
    const formattedDate = toMMDDYYYY(formData.date);
    const portCode = (formData.portCode || "").trim();
    const arrivalAirport = portCode; // same for all as you requested
    const locationOfGoods = LOCATION_MAP[arrivalAirport?.trim()] || ""; // derived

    const uploadRows = [];

    for (let i = 0; i < itemRows.length; i++) {
        const r = itemRows[i];
        const groupId = Math.floor(i / 998) + 1;

        // Start an empty row with all headers
        const newRow = Object.fromEntries(HEADERS.map((h) => [h, ""]));

        // Hardcodes
        Object.assign(newRow, HARDCODE);

        // Derived / combine / input
        newRow.GroupIdentifier = groupId;
        newRow.DescOfMerchandise = r[colCommodity];
        newRow.Description = r[colCommodity];
        newRow.HTS = r[colHTS];
        newRow.HTSQty = r[colQTY];
        newRow["Manifest Qty Piece count"] = r[colQTY];

        // HTSValue with floor rule
        newRow.HTSValue = parseFloat(r[colSubtotal]) || 0;

        newRow.InvoiceNumber = invoiceNumber;

        newRow.ManufacturerName = r[colManufName];
        newRow.ManufacturerStreetAddress = r[colManufAddr];
        newRow.ManufacturerCity = r[colManufCity];
        newRow.ManufacturerCountry = r[colManufCountry];
        newRow.ManufacturerPostalCode = r[colManufPostal];
        newRow.ManufacturerProvince = "";

        newRow["Airline 3 digit code"] = airlineCode;
        newRow["Master Bill Number"] = masterBillNumber;
        newRow["House AWB"] = formData.houseAWB || "";

        // Dates
        newRow["Arrival Date"] = formattedDate;
        newRow["Date of Export"] = formattedDate;
        newRow.EntryDate = formattedDate;
        newRow.ImportDate = formattedDate;

        // Ports / etc from user
        newRow.UnladingPort = portCode;
        newRow["Remote Port"] = portCode;
        newRow["STATE OF DESTINATION"] = formData.destinationState || "";
        newRow["Carrier Code"] = formData.carrierCode || "";
        newRow["Voyage Flight No"] = formData.voyageFlightNo || "";
        newRow["Arrival Airport"] = arrivalAirport;
        newRow["Location of Goods"] = locationOfGoods;

        // Countries
        const origin = r[colCountryOrigin];
        newRow["Country Of Origin"] = origin;
        newRow["Country of Export"] = origin;

        // Unit already hardcoded
        // Selling fields already hardcoded
        // Preparer Port already hardcoded

        uploadRows.push(newRow);
    }

    // Build workbook
    const sheet = xlsx.utils.json_to_sheet(uploadRows, { header: HEADERS });
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, sheet, "Sheet1");
    const buffer = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });

    return { buffer, airlineCode, masterBillNumber };
}

module.exports = { convertTemuExcel };
