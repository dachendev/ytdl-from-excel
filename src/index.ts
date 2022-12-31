import xlsx from "xlsx";
import fs from "node:fs";
import ytdl from "ytdl-core";
import path from "node:path";
import dotenv from "dotenv";

interface IIndexableObject {
  [key: string]: any;
}

// Load environment variables
// -----------------------------------------------------------------------------
dotenv.config();

if (!process.env.FILE_PATH) {
  throw new Error(`Missing required environment variable: "FILE_PATH".`);
}

if (!process.env.URL_COLUMN_NAME) {
  throw new Error(`Missing required environment variable: "URL_COLUMN_NAME".`);
}

const FILE_PATH = path.normalize(process.env.FILE_PATH);
const SHEET_NAME = process.env.SHEET_NAME ?? "Sheet1";
const URL_COLUMN_NAME = process.env.URL_COLUMN_NAME;

async function main() {
  process.stdout.write("Getting things ready...");

  // Read workbook
  // ---------------------------------------------------------------------------
  if (!fs.existsSync(FILE_PATH)) {
    throw new Error(`Cannot find file: "${FILE_PATH}".`);
  }

  const workbook = xlsx.readFile(FILE_PATH);
  
  if (!workbook.SheetNames.includes(SHEET_NAME)) {
    throw new Error(`Cannot find sheet: "${SHEET_NAME}".`);
  }

  const rows = xlsx.utils.sheet_to_json<IIndexableObject>(workbook.Sheets[SHEET_NAME]);

  if (!rows[0].hasOwnProperty(URL_COLUMN_NAME)) {
    throw new Error(`Cannot find column: "${URL_COLUMN_NAME}".`);
  }

  process.stdout.write("Done!\n");

  // Create output directory
  // ---------------------------------------------------------------------------
  const outDir = `out_${Date.now()}`;

  process.stdout.write(`Creating output directory "${outDir}"...`);

  await fs.promises.mkdir(outDir);

  process.stdout.write("Done!\n");

  // Download files
  // ---------------------------------------------------------------------------
  let index = 0;
  const count = rows.length;

  for (const row of rows) {
    const url = String(row[URL_COLUMN_NAME]);
  
    // Skip nullish values
    if (!url) continue;

    process.stdout.write(`Getting info from youtube.com (${url})...`);

    const { videoDetails } = await ytdl.getBasicInfo(url);

    process.stdout.write("Done!\n");
    
    const filename = `${videoDetails.author.name} - ${videoDetails.title}.mp3`;

    process.stdout.write(`Downloading "${filename}" [${index + 1}/${count}]...`);

    // Use promise to ensure async
    await new Promise<void>((resolve, reject) => {
      ytdl(url, { filter: "audioonly" })
        .pipe(fs.createWriteStream(path.join("./", outDir, filename)))
        .on("close", () => {
          process.stdout.write("Done!\n");
          resolve();
        });
    });

    index++;
  }
}

main();