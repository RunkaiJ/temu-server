const express = require("express");
const router = express.Router();
const { convertTemuExcel } = require("../utils/convertTemu");

router.post("/", async (req, res) => {
    console.log("📥 Received POST /api/convert");
    console.log("🗂️ req.files:", req.files);
    console.log("📝 req.body.form:", req.body.form);

    try {
        console.log("🔥 Incoming /api/convert request");
        // Validate files
        if (!req.files || !req.files.template || !req.files.combine) {
            console.log("Missing files:", req.files);
            return res
                .status(400)
                .json({ error: "Missing template or combine file." });
        }

        // Validate and parse form
        if (!req.body.form) {
            console.log("Missing form data");
            return res.status(400).json({ error: "Missing form data." });
        }

        let formData;

        try {
            formData = JSON.parse(req.body.form);
        } catch (parseError) {
            console.error("❌ JSON parse failed:", parseError);
            return res.status(400).json({ error: "Invalid form data" });
        }

        // Manual field checks
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
