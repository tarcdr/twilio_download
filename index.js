const fs = require('fs-extra');
const xlsx = require('xlsx');
const path = require('path');
const { Twilio } = require('twilio');
const { default: axios } = require('axios');
require('dotenv').config(); // โหลดค่าจากไฟล์ .env

// อ่านค่า environment variables
const accountSid = process.env.TWILIO_ACCOUNT_SID;
const authToken = process.env.TWILIO_AUTH_TOKEN;

// ตรวจสอบว่าค่า env ถูกต้องหรือไม่
if (!accountSid || !authToken) {
  throw new Error('Please provide valid TWILIO_ACCOUNT_SID and TWILIO_AUTH_TOKEN in .env file');
}

const client = new Twilio(accountSid, authToken);

// อ่านไฟล์ .xlsx ที่เก็บ refCode
const readExcelFile = (filePath) => {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(worksheet);
};

// เขียนผลลัพธ์กลับไปที่ไฟล์ .xlsx
const writeToExcelFile = (filePath, data) => {
  const newSheet = xlsx.utils.json_to_sheet(data);
  const newWorkbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(newWorkbook, newSheet);
  xlsx.writeFile(newWorkbook, filePath);
};

// Function to download a file
async function downloadFile(url, filePath) {
  const writer = fs.createWriteStream(filePath);
  const response = await axios({
    url,
    method: 'GET',
    responseType: 'stream',
    auth: {
      username: accountSid,
      password: authToken,
    },
  });

  response.data.pipe(writer);
  return new Promise((resolve, reject) => {
    writer.on('finish', resolve);
    writer.on('error', reject);
  });
}

// ฟังก์ชันหลักสำหรับการดาวน์โหลดไฟล์
const downloadRecording = async (refCode, outputDir) => {
  console.log(`Processing ... (${refCode})`);
  try {
    const rooms = await client.video.rooms.list({
      status: 'completed', // Fetch only completed rooms
      uniqueName: refCode
    });
    if (rooms.length === 0) {
      console.log(`No rooms found for refCode: ${refCode}`);
      return 'no room';
    }

    const downloaded = [];
    for (const room of rooms) {
      console.log(`Found room: ${room.sid}`);
      const recordings = await client.video.recordings.list({ groupingSid: room.sid });

      for (const recording of recordings) {
        const filePath = path.join(outputDir, `${recording.sid}.${recording.containerFormat}`);

        // Check if the file already exists
        if (fs.existsSync(filePath)) {
          console.log(`File already exists, skipping: ${filePath}`);
          downloaded.push(recording.sid);
          continue;  // Skip downloading this file if it already exists
        }

        const mediaUrl = `https://video.twilio.com/v1/Recordings/${recording.sid}/Media`;

        console.log(`Downloading recording ${recording.sid} (${recording.containerFormat})...`);
        await downloadFile(mediaUrl, filePath);
        downloaded.push(recording.sid);
        console.log(`Downloaded recording ${recording.sid} to ${filePath}`);
      }
      if (downloaded.length === 0) {
        console.log(`No files found for refCode: ${refCode}`);
        await fs.remove(outputDir); // ลบไฟล์ที่เสียหายออก
        console.log(`Removed corrupted files for refCode: ${refCode}`);
        return 'no files';
      }
    }
    return `success (${downloaded.length})`;
  } catch (error) {
    console.error(`Error downloading for refCode: ${refCode} - ${error.message}`);
    return 'error';
  }
};

// ฟังก์ชันหลักสำหรับการจัดการ refCode และดาวน์โหลดไฟล์
const processRefCodes = async (excelFilePath, outputDir) => {
  const refCodes = readExcelFile(excelFilePath);
  console.log(`read file ${excelFilePath}`);
  let timer = null;
  const results = {};
  for (const { refCode, status } of refCodes) {
    const result = { refCode, status };
    results[refCode] = result;
  }

  for (const { refCode, status } of refCodes) {
    const result = { refCode, status };
    (status) && console.log(`read ... ${refCode} status = "${status}"`, status.search('success') < 0);
    if (!status || status.search('success') < 0) {
      const refCodeOutputDir = path.join(outputDir, refCode);
      await fs.ensureDir(refCodeOutputDir);

      const newStatus = await downloadRecording(refCode, refCodeOutputDir);

      if (newStatus === 'error') {
        await fs.remove(refCodeOutputDir); // ลบไฟล์ที่เสียหายออก
        console.log(`Removed corrupted files for refCode: ${refCode}`);
      } else if (newStatus === 'no room') {
        await fs.remove(refCodeOutputDir); // // ลบ Folder ที่ Error ออก
      }
      result.status = newStatus;
    }
    results[refCode] = result;
    clearTimeout(timer);
    timer = setTimeout(() => {
      // เขียนผลลัพธ์กลับไปที่ไฟล์ .xlsx แทนไฟล์เดิม
      console.log('Write to excel file');
      writeToExcelFile(excelFilePath, Object.values(results));
    }, 10000);
  }

  const errors = Object.values(results).filter(o => o.status === 'error');
  console.log('errors', errors);
  for (const { refCode, status } of errors) {
    const result = { refCode, status };

    const refCodeOutputDir = path.join(outputDir, refCode);
    await fs.ensureDir(refCodeOutputDir);

    const newStatus = await downloadRecording(refCode, refCodeOutputDir);

    if (newStatus === 'error') {
      await fs.remove(refCodeOutputDir); // ลบไฟล์ที่เสียหายออก
      console.log(`Removed corrupted files for refCode: ${refCode}`);
    }
    results[refCode] = result;
  }

  // เขียนผลลัพธ์กลับไปที่ไฟล์ .xlsx แทนไฟล์เดิม
  writeToExcelFile(excelFilePath, Object.values(results));
};

// เริ่มการทำงาน
(async () => {
  const excelFilePath = './resource/refcode.xlsx'; // เส้นทางไฟล์ Excel
  const outputDir = './recordings'; // เส้นทางโฟลเดอร์ที่ต้องการเก็บไฟล์

  await processRefCodes(excelFilePath, outputDir);
})();
