#!/usr/bin/env node
'use strict';
const zlib = require('zlib');
const fs   = require('fs');
const path = require('path');

// ── CRC32 ──────────────────────────────────────────────────────────────────
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    t[i] = c;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = (c >>> 8) ^ CRC_TABLE[(c ^ buf[i]) & 0xff];
  return (c ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data) {
  const t   = Buffer.from(type, 'ascii');
  const len = Buffer.allocUnsafe(4); len.writeUInt32BE(data.length);
  const crc = Buffer.allocUnsafe(4); crc.writeUInt32BE(crc32(Buffer.concat([t, data])));
  return Buffer.concat([len, t, data, crc]);
}

// ── PNG builder ─────────────────────────────────────────────────────────────
function buildPNG(size, rgba) {
  const sig  = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  const ihdr = Buffer.allocUnsafe(13);
  ihdr.writeUInt32BE(size, 0); ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; ihdr[9] = 6; // 8-bit RGBA

  const rowLen = 1 + size * 4;
  const raw = Buffer.allocUnsafe(size * rowLen);
  for (let y = 0; y < size; y++) {
    raw[y * rowLen] = 0;
    raw.set(rgba.subarray(y * size * 4, (y + 1) * size * 4), y * rowLen + 1);
  }

  return Buffer.concat([
    sig,
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', zlib.deflateSync(raw, { level: 9 })),
    pngChunk('IEND', Buffer.alloc(0)),
  ]);
}

// ── 5×7 pixel font glyphs (bit4 = leftmost column) ──────────────────────────
const GLYPHS = {
  J: [0b01111, 0b00010, 0b00010, 0b00010, 0b10010, 0b11110, 0b01100],
  T: [0b11111, 0b00100, 0b00100, 0b00100, 0b00100, 0b00100, 0b00100],
};
const GW = 5, GH = 7;

// ── Draw icon at given size ───────────────────────────────────────────────────
function drawIcon(size) {
  const rgba = new Uint8Array(size * size * 4);

  // Fill background: #6366f1
  for (let i = 0; i < size * size; i++) {
    rgba[i*4]   = 0x63;
    rgba[i*4+1] = 0x66;
    rgba[i*4+2] = 0xf1;
    rgba[i*4+3] = 255;
  }

  // Rounded corners (~18% radius, like iOS)
  const r = Math.round(size * 0.18);
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      let outside = false;
      if (x < r && y < r)
        outside = (x-r)**2 + (y-r)**2 > r*r;
      else if (x >= size-r && y < r)
        outside = (x-(size-1-r))**2 + (y-r)**2 > r*r;
      else if (x < r && y >= size-r)
        outside = (x-r)**2 + (y-(size-1-r))**2 > r*r;
      else if (x >= size-r && y >= size-r)
        outside = (x-(size-1-r))**2 + (y-(size-1-r))**2 > r*r;
      if (outside) rgba[(y * size + x) * 4 + 3] = 0;
    }
  }

  // "JT" text: J then 2-unit gap then T
  const GAP   = 2;
  const total = GW * 2 + GAP;
  const scale = Math.max(1, Math.floor(size * 0.65 / total));
  const textW = total * scale;
  const textH = GH * scale;
  const ox    = Math.floor((size - textW) / 2);
  const oy    = Math.floor((size - textH) / 2);

  function drawGlyph(key, startX) {
    const rows = GLYPHS[key];
    for (let row = 0; row < GH; row++) {
      for (let col = 0; col < GW; col++) {
        if (rows[row] & (1 << (GW - 1 - col))) {
          for (let sy = 0; sy < scale; sy++) {
            for (let sx = 0; sx < scale; sx++) {
              const px = startX + col * scale + sx;
              const py = oy + row * scale + sy;
              if (px >= 0 && px < size && py >= 0 && py < size) {
                const i = (py * size + px) * 4;
                if (rgba[i+3] > 0) { // only on visible (non-corner) pixels
                  rgba[i] = rgba[i+1] = rgba[i+2] = 255;
                }
              }
            }
          }
        }
      }
    }
  }

  drawGlyph('J', ox);
  drawGlyph('T', ox + (GW + GAP) * scale);

  return rgba;
}

for (const size of [192, 512]) {
  const png  = buildPNG(size, drawIcon(size));
  const file = path.join(__dirname, `icon-${size}.png`);
  fs.writeFileSync(file, png);
  console.log(`✓  icon-${size}.png  (${png.length} bytes)`);
}
