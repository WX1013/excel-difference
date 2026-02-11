const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const toIco = require('to-ico');

async function build() {
  const src = path.resolve(__dirname, '..', 'src', 'images', 'logo.jpg');
  const outDir = path.resolve(__dirname, '..', 'public', 'favicons');
  if (!fs.existsSync(src)) {
    console.error('源图未找到:', src);
    process.exit(1);
  }
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const sizes = [16, 32, 48, 180];
  const pngBuffers = [];

  for (const size of sizes) {
    // 使用 SVG 圆形遮罩来裁剪为圆形（透明背景）
    const svgMask = `<svg width="${size}" height="${size}" xmlns="http://www.w3.org/2000/svg"><circle cx="${size/2}" cy="${size/2}" r="${size/2}" fill="#fff"/></svg>`;

    const buf = await sharp(src)
      .resize(size, size, { fit: 'cover' })
      .png()
      .composite([{ input: Buffer.from(svgMask), blend: 'dest-in' }])
      .toBuffer();

    const fileName = size === 180 ? `apple-touch-icon.png` : `icon-${size}.png`;
    fs.writeFileSync(path.join(outDir, fileName), buf);
    if (size === 16 || size === 32 || size === 48) pngBuffers.push(buf);
  }

  // 生成 favicon.ico（包含 16/32/48）
  const icoBuf = await toIco(pngBuffers);
  fs.writeFileSync(path.join(outDir, 'favicon.ico'), icoBuf);

  console.log('圆形图标生成完成，输出目录：', outDir);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
