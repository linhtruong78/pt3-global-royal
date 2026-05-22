import sharp from 'sharp';
import { readFileSync, writeFileSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const imgDir = join(__dirname, '../assets/images');

async function removeDarkBackground(inputFile, outputFile, threshold = 60) {
  const input = join(imgDir, inputFile);
  const output = join(imgDir, outputFile);

  const { data, info } = await sharp(input)
    .ensureAlpha()
    .raw()
    .toBuffer({ resolveWithObject: true });

  const pixels = new Uint8ClampedArray(data);
  const { width, height, channels } = info;

  for (let i = 0; i < pixels.length; i += channels) {
    const r = pixels[i];
    const g = pixels[i + 1];
    const b = pixels[i + 2];

    // Remove dark pixels (black or near-black or dark burgundy)
    const brightness = (r + g + b) / 3;
    const isDark = brightness < threshold && r < 100 && g < 60 && b < 80;

    if (isDark) {
      pixels[i + 3] = 0; // transparent
    }
  }

  await sharp(pixels, { raw: { width, height, channels } })
    .png()
    .toFile(output);

  console.log(`✓ Saved: ${outputFile}`);
}

// Logo v1 — pure black background
await removeDarkBackground('logo-v1.png.jpg', 'logo-v1-transparent.png', 55);

// Logo v2 — dark burgundy background (slightly higher threshold)
await removeDarkBackground('logo-v2.png.jpg', 'logo-v2-transparent.png', 70);

console.log('Done — both logos processed.');
