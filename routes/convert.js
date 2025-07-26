const express = require("express");
const router = express.Router();
const { convertTemuExcel } = require("../utils/convertTemu");

router.post("/", async (req, res) => {
    try {
        if (!req.files?.manifest) {
            return res.status(400).json({ error: "Manifest file is missing." });
        }
        if (!req.body.form) {
            return res.status(400).json({ error: "Form data is missing." });
        }

        const formData = JSON.parse(req.body.form);

        // perform conversion
        const { buffer } = await convertTemuExcel(
            req.files.manifest.data,
            formData
        );

        // build filename from user inputs
        const { airlineCode, masterBillNumber } = formData;
        const filename = `${airlineCode}${masterBillNumber} UPLOAD.xlsx`;

        res.set({
            "Content-Type":
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "Content-Disposition": `attachment; filename="${filename}"`,
        });

        res.send(buffer);
    } catch (err) {
        console.error("Conversion error:", err);
        res.status(500).json({ error: "Failed to process file." });
    }
});
module.exports = router;
