const express = require("express");
const cors = require("cors");
const fileUpload = require("express-fileupload");
const convertRoute = require("./routes/convert");

const app = express();
const PORT = 5000;

app.use(
    cors({
        origin: "https://alex-toolbox.vercel.app",
        credentials: true,
        exposedHeaders: ["Content-Disposition"],
    })
);

app.use(fileUpload());

app.use("/api/convert", convertRoute);

app.listen(PORT, () => {
    console.log(`🚀 Server running on http://localhost:${PORT}`);
});
