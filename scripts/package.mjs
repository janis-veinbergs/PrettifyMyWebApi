import fs from 'node:fs';
import path from 'node:path';
import archiver from 'archiver';

const rootDir = process.cwd();
const srcDir = path.join(rootDir, 'src');
const manifestPath = path.join(srcDir, 'manifest.json');

if (!fs.existsSync(manifestPath)) {
    throw new Error(`Cannot find manifest: ${manifestPath}`);
}

const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
const version = manifest.version;

if (!version || typeof version !== 'string') {
    throw new Error('manifest.json must contain a string "version"');
}

const outputName = `PrettifyMyWebApi-v${version}.zip`;
const outputPath = path.join(rootDir, outputName);

if (fs.existsSync(outputPath)) {
    fs.unlinkSync(outputPath);
}

const output = fs.createWriteStream(outputPath);
const archive = archiver('zip', { zlib: { level: 9 } });

const done = new Promise((resolve, reject) => {
    output.on('close', resolve);
    output.on('error', reject);
    archive.on('error', reject);
});

archive.pipe(output);
archive.glob('**/*', {
    cwd: srcDir,
    nodir: true,
    dot: false,
    ignore: ['**/.DS_Store', 'prettifyWebApi.zip']
});

await archive.finalize();
await done;

console.log(`Created ${outputName}`);
