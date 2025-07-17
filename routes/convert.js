const express = require("express");
const router = express.Router();
const { convertTemuExcel } = require("../utils/convertTemu");

router.post("/", async (req, res) => {
    try {
        if (!req.files || !req.files.template || !req.files.combine) {
            return res.status(400).json({ error: "Files are missing." });
        }

        let formData;
        try {
            formData =
                typeof req.body.form === "string"
                    ? JSON.parse(req.body.form)
                    : req.body.form;
        } catch (err) {
            return res.status(400).json({ error: "Invalid form data format." });
        }

        // Manual field validation (adjust as needed)
        const requiredFields = [
            "date",
            "unladingPort",
            "arrivalAirport",
            "preparerPort",
            "remotePort",
            "destinationState",
            "locationOfGoods",
            "carrierCode",
            "voyageFlightNo",
            "houseAWB",
        ];

        for (const field of requiredFields) {
            if (!formData[field]) {
                return res
                    .status(400)
                    .json({ error: `Missing field: ${field}` });
            }
        }

        const { buffer, airlineCode, masterBillNumber } =
            await convertTemuExcel(
                req.files.template.data,
                req.files.combine.data,
                formData
            );

        const filename = `${airlineCode}${masterBillNumber} UPLOAD GENERATED.xlsx`;

        res.set({
            "Content-Type":
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "Content-Disposition": `attachment; filename="${filename}"`,
        });

        return res.send(buffer);
    } catch (err) {
        console.error("Conversion error:", err);
        res.status(500).json({ error: "Failed to process file." });
    }
});

module.exports = router;
