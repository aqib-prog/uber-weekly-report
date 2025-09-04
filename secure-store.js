// secure-store.js - File-based encryption without keytar
const crypto = require("crypto");
const os = require("os");
const path = require("path");
const fs = require("fs").promises;

const SERVICE = "UberWeeklyReporter";

// Generate a consistent key based on machine info
function generateMachineKey() {
  const machineId = os.hostname() + os.platform() + os.arch();
  return crypto
    .createHash("sha256")
    .update(machineId + SERVICE)
    .digest();
}

function sessionPath() {
  const dir = path.join(os.homedir(), ".uber-weekly-reporter");
  return path.join(dir, "session.enc");
}

async function encryptToFile(plaintextBuffer) {
  const key = generateMachineKey();
  const file = sessionPath();
  await fs.mkdir(path.dirname(file), { recursive: true });

  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  const enc = Buffer.concat([cipher.update(plaintextBuffer), cipher.final()]);
  const tag = cipher.getAuthTag();

  await fs.writeFile(file, Buffer.concat([iv, tag, enc]));
  return file;
}

async function decryptFromFile() {
  const key = generateMachineKey();
  const file = sessionPath();
  const raw = await fs.readFile(file);
  const iv = raw.subarray(0, 12);
  const tag = raw.subarray(12, 28);
  const enc = raw.subarray(28);

  const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(enc), decipher.final()]);
}

module.exports = {
  encryptToFile,
  decryptFromFile,
  sessionPath,
};
