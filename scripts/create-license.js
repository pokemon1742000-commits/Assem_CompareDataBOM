const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');

const count = Number.parseInt(process.argv[2] || '10', 10);
const outputPath = path.resolve(process.argv[3] || 'licenses.json');

if (!Number.isInteger(count) || count <= 0) {
  console.error('Usage: node scripts/create-license.js <count> [output-file]');
  process.exit(1);
}

const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

function createSegment(length) {
  let segment = '';
  for (let index = 0; index < length; index += 1) {
    const byte = crypto.randomBytes(1)[0];
    segment += alphabet[byte % alphabet.length];
  }
  return segment;
}

function createLicense() {
  return `${createSegment(3)}-${createSegment(3)}-${createSegment(3)}-${createSegment(4)}`;
}

const existing = readExistingLicenses(outputPath);
const licenses = new Set(existing);

while (licenses.size < existing.length + count) {
  licenses.add(createLicense());
}

fs.writeFileSync(outputPath, JSON.stringify({ licenses: Array.from(licenses) }, null, 2), 'utf8');
console.log(`Created ${count} license(s) in ${outputPath}`);

function readExistingLicenses(filePath) {
  try {
    const raw = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    if (Array.isArray(raw)) return raw;
    if (Array.isArray(raw.licenses)) return raw.licenses;
  } catch {
    return [];
  }

  return [];
}
