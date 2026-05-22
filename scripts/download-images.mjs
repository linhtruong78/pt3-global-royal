/**
 * Downloads and optimizes stock photos from Unsplash for the PT3 site.
 * Run: node scripts/download-images.mjs
 * Requires: sharp (already installed as devDependency)
 */

import sharp from 'sharp';
import { writeFile, mkdir } from 'fs/promises';
import { existsSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUTPUT_DIR = join(__dirname, '..', 'assets', 'images');

// Unsplash photo IDs + target specs
const IMAGES = [
  {
    id: 'photo-1510812431401-41d2bd2722f3',
    out: 'wine-collection.jpg',
    width: 900,
    height: 1200,
    quality: 78,
    label: 'Wine Collection Card',
  },
  {
    id: 'photo-1474722883778-792e7990302f',
    out: 'wine-hero.jpg',
    width: 1600,
    height: 900,
    quality: 80,
    label: 'Reserve Wines Hero',
  },
  {
    id: 'photo-1535958636474-b021ee887b13',
    out: 'beer-collection.jpg',
    width: 900,
    height: 1200,
    quality: 78,
    label: 'Beer Collection Card',
  },
  {
    id: 'photo-1558618666-fcd25c85cd64',
    out: 'beer-hero.jpg',
    width: 1600,
    height: 900,
    quality: 80,
    label: 'Shams Beer Hero',
  },
  {
    id: 'photo-1547592166-23ac45744acd',
    out: 'heritage-visual.jpg',
    width: 900,
    height: 1100,
    quality: 78,
    label: 'Heritage Visual Panel',
  },
  {
    id: 'photo-1565343116253-b04e34e3f9a6',
    out: 'hero-bg.jpg',
    width: 1800,
    height: 1000,
    quality: 72,
    label: 'Homepage Hero Background',
  },
];

async function download(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status} — ${url}`);
  const buf = await res.arrayBuffer();
  return Buffer.from(buf);
}

async function run() {
  if (!existsSync(OUTPUT_DIR)) {
    await mkdir(OUTPUT_DIR, { recursive: true });
  }

  for (const img of IMAGES) {
    const url =
      `https://images.unsplash.com/${img.id}` +
      `?w=${img.width}&h=${img.height}&q=90&auto=format&fit=crop`;

    const outPath = join(OUTPUT_DIR, img.out);
    process.stdout.write(`⬇  ${img.label} … `);

    try {
      const raw = await download(url);

      await sharp(raw)
        .resize(img.width, img.height, { fit: 'cover', position: 'centre' })
        .jpeg({ quality: img.quality, mozjpeg: true, progressive: true })
        .toFile(outPath);

      const stats = await import('fs').then(m => m.statSync(outPath));
      const kb = Math.round(stats.size / 1024);
      console.log(`✓  saved ${img.out} (${kb} KB)`);
    } catch (err) {
      console.error(`✗  FAILED — ${err.message}`);
    }
  }

  console.log('\nAll done! Check assets/images/');
}

run();
