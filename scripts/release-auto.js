const { execFileSync } = require('node:child_process');
const fs = require('node:fs');
const packageJson = require('../package.json');

const [requestedVersion, requestedMessage] = process.argv.slice(2);

function parseVersion(value) {
  const match = String(value || '').match(/^v?(\d+)\.(\d+)\.(\d+)$/);
  return match ? { major: Number(match[1]), minor: Number(match[2]), patch: Number(match[3]) } : null;
}

function nextVersion() {
  let latest = parseVersion(packageJson.version) || { major: 0, minor: 0, patch: 0 };
  try {
    const output = execFileSync('gh.exe', [
      'release', 'list', '--repo', 'pokemon1742000-commits/Assem_CompareDataBOM',
      '--limit', '20', '--json', 'tagName,isDraft',
      '--jq', '[.[] | select(.isDraft == false)][0].tagName'
    ], { encoding: 'utf8', windowsHide: true }).trim();
    latest = parseVersion(output) || latest;
  } catch {
    // Dùng version trong package.json nếu chưa lấy được latest từ GitHub.
  }
  return `${latest.major}.${latest.minor}.${latest.patch + 1}`;
}

const version = requestedVersion || nextVersion();
const releaseMessage = requestedMessage || `Release v${version}`;
const repository = 'pokemon1742000-commits/Assem_CompareDataBOM';
const tag = `v${version}`;

if (!version || !/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$/.test(version)) {
  console.error('Cach dung: npm run release:auto [-- version "Release message"]');
  process.exit(1);
}

function run(command, args, env = {}) {
  console.log(`\n> ${command} ${args.join(' ')}`);
  const executable = process.platform === 'win32' && command === 'npm' ? 'npm.cmd' : command;
  execFileSync(executable, args, {
    stdio: 'inherit',
    windowsHide: false,
    shell: process.platform === 'win32' && command === 'npm',
    env: { ...process.env, ...env }
  });
}

function getGithubToken() {
  if (process.env.GH_TOKEN) return process.env.GH_TOKEN;

  try {
    return execFileSync(process.platform === 'win32' ? 'gh.exe' : 'gh', ['auth', 'token'], {
      encoding: 'utf8',
      windowsHide: true
    }).trim();
  } catch {
    throw new Error('Chua dang nhap GitHub CLI. Hay chay: gh auth login');
  }
}

if (packageJson.version !== version) {
  run('npm', ['version', version, '--no-git-tag-version']);
} else {
  console.log(`Version ${version} da duoc cap nhat, bo qua buoc npm version.`);
}
fs.writeFileSync('update.json', `${JSON.stringify({
  version,
  url: `https://github.com/${repository}/releases/download/${tag}/Inventory-Compare-Setup-v${version}-x64.exe`
}, null, 2)}\n`, 'utf8');
const githubToken = getGithubToken();
run('npm', ['run', 'build:win', '--', '--publish', 'never']);
run('git', ['add', '-A']);

try {
  run('git', ['commit', '-m', releaseMessage]);
} catch (error) {
  console.error('Khong the commit thay doi. Qua trinh phat hanh da dung lai.');
  process.exit(error.status || 1);
}

run('git', ['push', 'origin', 'HEAD']);

const releaseAssets = [
  `release/Inventory-Compare-Setup-v${version}-x64.exe`,
  `release/Inventory-Compare-Setup-v${version}-x64.exe.blockmap`,
  'release/latest.yml',
  'update.json'
];
run('gh', [
  'release', 'create', tag, ...releaseAssets,
  '--repo', repository,
  '--title', tag,
  '--notes', releaseMessage,
  '--latest'
], { GH_TOKEN: githubToken });
console.log(`\nDa phat hanh ${tag} voi mot release duy nhat len GitHub.`);