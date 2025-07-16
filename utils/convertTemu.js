const xlsx = require("xlsx");

async function convertTemuExcel(templateBuffer, combineBuffer, formData) {
    const templateWorkbook = xlsx.read(templateBuffer, { type: "buffer" });
    const combineWorkbook = xlsx.read(combineBuffer, { type: "buffer" });

    const templateSheet = templateWorkbook.Sheets["Sheet1"];
    const combineSheet = combineWorkbook.Sheets[combineWorkbook.SheetNames[0]];

    const templateData = xlsx.utils.sheet_to_json(templateSheet, {
        defval: "",
    });
    const templateRaw = xlsx.utils.sheet_to_json(templateSheet, { header: 1 });
    const combineRaw = xlsx.utils.sheet_to_json(combineSheet, {
        header: 1,
        defval: "",
    });

    const headers = combineRaw[9];
    const itemRows = combineRaw.slice(10).filter((r) => r[2] && r[3]);

    const normalize = (str) =>
        str
            ?.toString()
            .replace(/\u00A0/g, " ")
            .trim()
            .toLowerCase();
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

    let invoiceNumber = "";
    for (const row of combineRaw) {
        const idx = row.findIndex(
            (cell) =>
                cell && cell.toString().toLowerCase().includes("invoice number")
        );
        if (idx !== -1 && row[idx + 1]) {
            invoiceNumber = row[idx + 1].toString().trim();
            break;
        }
    }

    let airlineCode = "",
        masterBillNumber = "";
    for (const row of combineRaw) {
        const idx = row.findIndex(
            (cell) => cell && cell.toString().toLowerCase().includes("po no")
        );
        if (idx !== -1 && row[idx + 1]) {
            const match = row[idx + 1].toString().match(/_(\d+)-(\d+)_/);
            if (match) {
                airlineCode = match[1];
                masterBillNumber = match[2];
            }
            break;
        }
    }

    const baseRow = templateData[0];
    const originalColumnOrder = templateRaw[0];

    const uploadRows = [];

    for (let i = 0; i < itemRows.length; i++) {
        const row = itemRows[i];
        const groupId = Math.floor(i / 998) + 1;
        const newRow = { ...baseRow };

        newRow.EntryType = "01";
        newRow.GroupIdentifier = groupId;
        newRow.ForwardingID = baseRow.ForwardingID;
        newRow.ConsigneeId = baseRow.ConsigneeId;
        newRow.RelatedParties = "Y";

        newRow.DescOfMerchandise = row[2];
        newRow.Description = row[2];
        newRow.InvoiceNumber = invoiceNumber;
        newRow.HTS = row[colHTS];
        newRow.HTSQty = row[colQTY];
        newRow.HTSValue = parseFloat(row[colSubtotal]) || 0;
        newRow["Manifest Qty Piece count"] = row[colQTY];

        newRow.ManufacturerName = row[colManufName];
        newRow.ManufacturerStreetAddress = row[colManufAddr];
        newRow.ManufacturerCity = row[colManufCity];
        newRow.ManufacturerProvince = "";
        newRow.ManufacturerCountry = row[colManufCountry];
        newRow.ManufacturerPostalCode = row[colManufPostal];

        newRow.UnladingPort = formData.unladingPort || "";
        newRow["Arrival Airport"] = formData.arrivalAirport || "";
        newRow["Preparer Port"] = formData.preparerPort || "";
        newRow["Remote Port"] = formData.remotePort || "";
        newRow["STATE OF DESTINATION"] = formData.destinationState || "";
        newRow["Location of Goods"] = formData.locationOfGoods || "";
        newRow["Carrier Code"] = formData.carrierCode || "";
        newRow["Voyage Flight No"] = formData.voyageFlightNo || "";
        newRow["Arrival Date"] = formData.date || "";
        newRow["Date of Export"] = formData.date || "";
        newRow.EntryDate = formData.date || "";
        newRow.ImportDate = formData.date || "";
        newRow["Airline 3 digit code"] = airlineCode;
        newRow["Master Bill Number"] = masterBillNumber;
        newRow["House AWB"] = formData.houseAWB || "";

        newRow.SellingName = baseRow.SellingName;
        newRow.SellingStreetAddress = baseRow.SellingStreetAddress;
        newRow.SellingState = "";
        newRow.SellingPostalCode = baseRow.SellingPostalCode;
        newRow.Unit = "PCS";

        uploadRows.push(newRow);
    }

    const sheet = xlsx.utils.json_to_sheet(uploadRows, {
        header: originalColumnOrder,
    });
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, sheet, "Sheet1");

    const outputBuffer = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });

    return {
        buffer: outputBuffer,
        airlineCode,
        masterBillNumber,
    };
}

module.exports = { convertTemuExcel };
