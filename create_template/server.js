const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
const fs = require("fs").promises;
const path = require("path");
const os = require("os");

async function generateTemplates() {
    try {
        const jsonDir = path.resolve(__dirname, "JsonData");
        const jsonFiles = await fs.readdir(jsonDir);

        // Determine how many workers (PM2 instances) are running
        const totalWorkers = os.cpus().length; // Matches PM2 -i max
        const workerId = parseInt(process.env.NODE_APP_INSTANCE) || 0;

        // Split JSON files among workers
        const chunkSize = Math.ceil(jsonFiles.length / totalWorkers);
        const filesForThisWorker = jsonFiles.slice(workerId * chunkSize, (workerId + 1) * chunkSize);

        console.log(`Worker ${workerId} processing ${filesForThisWorker.length} files...`);

        // Load the docx template as binary content once
        const content = await fs.readFile(path.resolve(__dirname, "125418pg1.docx"), "binary");

        for (const jsonFile of filesForThisWorker) {
            const jsonData = JSON.parse(await fs.readFile(path.join(jsonDir, jsonFile), "utf8"));
            const dataToRender = Array.isArray(jsonData) ? jsonData[0] : jsonData;
            if (Array.isArray(dataToRender.Items_to_be_purchased)) {
                dataToRender.Items_to_be_purchased = dataToRender.Items_to_be_purchased.flat();
            }

            const zip = new PizZip(content);
            const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

            doc.render(dataToRender);

            const buf = doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });

            const outputFilePath = path.resolve(__dirname, `output_${path.basename(jsonFile, ".json")}.docx`);
            await fs.writeFile(outputFilePath, buf);
            console.log(`✅ Worker ${workerId} generated: ${outputFilePath}`);
        }
    } catch (error) {
        console.error(`❌ Worker ${process.env.NODE_APP_INSTANCE} error:`, error);
    }
}

// Run the function
generateTemplates();

// pm2 stop server.js
// pm2 start server.js -i max
