const { createCanvas, loadImage } = require('canvas');
const fs = require('fs');

// 创建canvas实例
const canvas = createCanvas(760, 760);
const ctx = canvas.getContext('2d');

// 加载背景图片
loadImage('1DB-EBDW1.png').then(bgImg => {
  // 加载需要合并的图片
  loadImage('Como Ash 01 1S.png').then(fgImg => {
    // 先绘制背景图片
    ctx.drawImage(bgImg, 0, 0, canvas.width, canvas.height);

    // 计算需要合并的图片的缩放比例
    const scale = 0.4;
    const fgWidth = fgImg.width * scale;
    const fgHeight = fgImg.height * scale;

    // 将需要合并的图片绘制在背景图片的上一层，并将图像改成圆形
    ctx.save();
    ctx.beginPath();
    ctx.arc(canvas.width - fgWidth / 2, fgHeight / 2, fgWidth / 2, 0, Math.PI * 2, true);
    ctx.closePath();
    ctx.clip();
    ctx.drawImage(fgImg, canvas.width - fgWidth, 0, fgWidth, fgHeight);
    ctx.restore();

    // 将合并后的图片保存到本地
    const mergedImg = canvas.toBuffer();
    fs.writeFileSync('merged.png', mergedImg);
  });
});
