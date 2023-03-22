const { createCanvas, loadImage } = require('canvas');
const fs = require('fs');

// 创建canvas实例
const canvas = createCanvas(450, 450); //画布的大小 
const ctx = canvas.getContext('2d');

// 读取背景图片文件夹
const bgImgDir = '/Users/ryanweng/Documents/Cuppowood/website/产品导入/Shopify/Accessorise/';
const bgImgs = fs.readdirSync(bgImgDir);

// 读取需要合并的图片文件夹
const fgImgDir = '/Users/ryanweng/Documents/Cuppowood/website/产品导入/Shopify/Colors/';
const fgImgs = fs.readdirSync(fgImgDir);


const outputDir = '/Users/ryanweng/Documents/Cuppowood/website/产品导入/Shopify/AccConcatPhoto/';
if(fs.existsSync(outputDir)){
  fs.rm(outputDir, { recursive: true });
}

if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir);
}

// 循环处理每组图片
for (let i = 0; i < bgImgs.length; i++) {
  const bgImgPath = bgImgDir + bgImgs[i];

  // 加载背景图片
  loadImage(bgImgPath).then(bgImg => {

    // 循环处理每张前景图片
    for (let j = 0; j < fgImgs.length; j++) {
      const fgImgPath = fgImgDir + fgImgs[j];

      // 加载需要合并的图片
      loadImage(fgImgPath).then(fgImg => {
        // 先绘制背景图片
        ctx.drawImage(bgImg, 0, 0, canvas.width, canvas.height);

        // 计算需要合并的图片的缩放比例
        const fgWidth = 120;
        const fgHeight = 120;

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

        // /\s+/g：匹配所有的空格字符，包括空格、制表符、换行符等。
        // /.[^/.]+$/：匹配文件名中的扩展名部分。其中，点号表示任意字符，除了换行符，
        // [^/.]+ 表示除了点号和斜杠之外的任意字符，+表示匹配一个或多个这样的字符，$表示匹配字符串结尾。

        const bgImgName = bgImgs[i].replace(/\s+/g, '').replace(/.[^/.]+$/, ''); 
        const fgImgName = fgImgs[j].replace(/\s+/g, '').replace(/-/g, '');
        const outputName = `${bgImgName}--${fgImgName}`; // 修改名字格式
        fs.writeFileSync(outputDir + outputName, mergedImg);
      }).catch(error => {
        console.error(`Error processing image ${fgImgPath}:`, error);
      });
    }
  });
}
