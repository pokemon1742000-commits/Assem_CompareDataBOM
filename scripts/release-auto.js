const { execFileSync } = require('node:child_process');
const packageJson = require('../package.json');

const [version, releaseMessage = `Release v${version || ''}`] = process.argv.slice(2);

if (!version || !/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$/.test(version)) {
  console.error('Cach dung: npm run release:auto -- 3.0.1 "Release v3.0.1"');
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
run('npm', ['run', 'build:win', '--', '--publish', 'always'], {
  GH_TOKEN: getGithubToken()
});
run('git', ['add', '-A']);

try {
  run('git', ['commit', '-m', releaseMessage]);
} catch (error) {
  console.error('Khong the commit thay doi. Qua trinh phat hanh da dung lai.');
  process.exit(error.status || 1);
}

run('git', ['push', 'origin', 'HEAD']);
console.log(`\nDa phat hanh v${version} len GitHub Releases.`);