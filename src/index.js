import pptxgen from "pptxgenjs";
import { parse } from 'csv-parse';
import { promises as fs } from 'fs'

const CENTIMETER_PER_INCH = 2.54;
const ONE_CENTIMETER = 1;
const TWO_CENTIMETERS = 2;
const ONE_CENTIMETER_IN_INCH = ONE_CENTIMETER / CENTIMETER_PER_INCH;
const TWO_CENTIMETER_IN_INCH = TWO_CENTIMETERS / CENTIMETER_PER_INCH;

const fontFace = 'Malgun Gothic';

const black = "000000";
const white = "FFFFFF";

let pres = new pptxgen();

function addNameCard(slide, position, size, tag) {
  const x = position.x;
  const y = position.y;
  const width = size.width;
  const height = size.height;

  const titleFontSize = 18;
  const graduateFontSize = 20;
  const nameFontSize = 35;

  slide.addText(`\n\n${tag.name}`, {
    shape: pres.ShapeType.rect,
    x: x, y: y,
    w: width, h: height,
    fontFace: fontFace,
    fontSize: nameFontSize,
    fill: {color: white},
    line: {color: black},
    align: 'center',
    bold: true
  });

  slide.addText(`\n${tag.title}`, {
    x: x, y: y + TWO_CENTIMETER_IN_INCH,
    w: width, h: ONE_CENTIMETER_IN_INCH,
    fontFace: fontFace,
    fontSize: titleFontSize,
    align: 'center',
    bold: true
  })

  if ( tag.graduate !== undefined && tag.graduate !== null ) {
    slide.addText(`${tag.graduate}회 (${tag.hometown})`, {
      x: x, y: y + ONE_CENTIMETER_IN_INCH,
      w: width, h: ONE_CENTIMETER_IN_INCH,
      fontFace: fontFace,
      fontSize: graduateFontSize,
      align: 'center',
      bold: true
    })
  }
}

function addNameCards(slide, position, size, tags)
{
  const cardPosition = { x: position.x, y: position.y }
  const cardWidth = size.width;
  const GAP_BETWEEN_CARDS_IN_INCH = 0.1;

  for ( const tag of tags) {
    addNameCard(slide, cardPosition, size, tag);
    cardPosition.x += cardWidth + GAP_BETWEEN_CARDS_IN_INCH;
  }
}

function generateNameCardSlideFor(visitors) {
  let slide = pres.addSlide();

  const EIGHT_CENTIMETER = 8;
  const SIX_CENTIMETER = 6;
  const width = EIGHT_CENTIMETER/CENTIMETER_PER_INCH;
  const height = SIX_CENTIMETER/CENTIMETER_PER_INCH;

  const CARD_MARGIN_IN_INCH = 0.1;

  const position = {
    x: CARD_MARGIN_IN_INCH,
    y: CARD_MARGIN_IN_INCH,
  };

  const size = {
    width,
    height
  }

  const CARD_COUNT_PER_ROW = 3;
  const CARD_HEIGHT = size.height;

  addNameCards(slide, position, size, visitors.slice(0,CARD_COUNT_PER_ROW));

  position.x = CARD_MARGIN_IN_INCH;
  position.y = position.y + CARD_HEIGHT + CARD_MARGIN_IN_INCH;
  addNameCards(slide, position, size, visitors.slice(CARD_COUNT_PER_ROW,SIX_CENTIMETER));
}

function generateNameCardPPTfor(visitors) {
  const CARD_COUNT_PER_SLIDE = 6;
  var visitorIndex = 0;

  while ( visitorIndex < visitors.length ) {
    const visitorsInPage = visitors.slice(visitorIndex, visitorIndex + CARD_COUNT_PER_SLIDE);
    generateNameCardSlideFor(visitorsInPage);
    visitorIndex += CARD_COUNT_PER_SLIDE;
  }
}

function createVisitorsFrom(records) {
  const visitors = records.map((record) => {
    return {
      name : record[0],
      title: record[1],
      graduate: record[2],
      hometown: record[3]
    }
  });

  return visitors;
}

async function generateNameCardPPTfrom(csvInputFilename, pptOutputFilename) {
  const content = await fs.readFile(csvInputFilename);

  parse(content, { delimiter: ','}, function(err, records) {
    const visitors = createVisitorsFrom(records);
    generateNameCardPPTfor(visitors);
    pres.writeFile({ fileName: pptOutputFilename });
  });
}

if ( process.argv.length !== 4 ) {
  console.error('Usage: node index.js csv-input-file ppt-output-file');
  console.error('Read README.md for the detail.');
  process.exit(1);
}

generateNameCardPPTfrom(process.argv[2], process.argv[3]).then(() => {});
