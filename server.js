const pptxgen = require('pptxgenjs');
const got = require('got');
const express = require('express');
require('dotenv').config()
const app = express();


const redisClient = require('./redis-client');

const generatePowerpoint = (text) => {
  // 1. Create a Presentation
  const pres = new pptxgen();

  // 2. Add a Slide to the presentation
  const slide = pres.addSlide();

  // 3. Add 1+ objects (Tables, Shapes, etc.) to the Slide
  slide.addText(text ?? 'Test String', {
    x: 1.5,
    y: 1.5,
    color: '363636',
    fill: { color: 'F1F1F1' },
    align: pres.AlignH.center,
  });

  // 4. Save the Presentation
  return pres.stream();
};

const getPowerPoint = async (req, res, next) => {
  try {
    const pptxData = await getPptxData();
    const { title, datasets } = pptxData[process.env.DATA_ID];
    const keyValue = datasets[0].data[1];
    const pptx = await generatePowerpoint(`${title}: ${keyValue['key']} - ${keyValue['value']}`);

    res.writeHead(200, { 'Content-disposition': 'attachment;filename=Test.pptx', 'Content-Length': pptx.length });
    res.end(Buffer.from(pptx, 'binary'));
  } catch (error) {
    next(error);
  }
};

const getPptxData = async () => {
  const cacheKey = 'pptx_data';
  const cacheValue = await redisClient.getAsync(cacheKey);
  if (!cacheValue) {
    console.log('FETCHING');
    const pptxData = await got.post(process.env.DATA_URL, {
      headers: {
        Authorization: `Bearer ${process.env.TOKEN}`,
      },
    });
    redisClient.setAsync(cacheKey, pptxData.body);

    return JSON.parse(pptxData.body);
  }
  console.log('USING CACHE');

  return JSON.parse(cacheValue);
};

app.get('/powerpoint', getPowerPoint);

app.get('/', (req, res) => {
  return res.send('Hello world');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
