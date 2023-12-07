const AppEnv = gsGetEnv();

function myFunction() {
  const tableId = AppEnv.price__google_sheet_id;
  const table = SpreadsheetApp.openById(tableId);

  const listBYN = table.getSheetByName('BYN');
  const listUSD = table.getSheetByName('USD');
  const listRUB = table.getSheetByName('RUB');
  const listGEL = table.getSheetByName('GEL');
  const listAMD = table.getSheetByName('AMD');

  listBYN.clear();
  listUSD.clear();
  listRUB.clear();
  listGEL.clear();
  listAMD.clear();

  const price_BYN_xlsx_array: string[][] = [];
  const price_USD_xlsx_array: string[][] = [];
  const price_RUB_xlsx_array: string[][] = [];
  const price_GEL_xlsx_array: string[][] = [];
  const price_AMD_xlsx_array: string[][] = [];

  const otherContacts = AppEnv.price__contactInHeader.split('\\n').join('\n');
  const date = DateController.getMyFullDate();
  const titles = {
    BYN: `Прайс от ${date} в белорусских рублях (BYN) (с доставкой и НДС по Беларуси)\n${otherContacts}`,
    USD: `Прайс от ${date} в долларах США (USD) (самовывоз из Турции)\n${otherContacts}`,
    RUB: `Прайс от ${date} в российских рублях (RUB) (с доставкой и НДС по России)\n${otherContacts}`,
    GEL: `Прайс от ${date} в грузинских лари (GEL) (с доставкой и НДС по Грузии)\n${otherContacts}`,
    AMD: `Прайс от ${date} в армянских драмах (AMD) (с доставкой и НДС по Армении)\n${otherContacts}`,
  };

  price_BYN_xlsx_array.push([titles['BYN'], '', '', '', '', '', '', '', '', '']);
  price_USD_xlsx_array.push([titles['USD'], '', '', '', '', '', '', '', '', '']);
  price_RUB_xlsx_array.push([titles['RUB'], '', '', '', '', '', '', '', '', '']);
  price_GEL_xlsx_array.push([titles['GEL'], '', '', '', '', '', '', '', '', '']);
  price_AMD_xlsx_array.push([titles['AMD'], '', '', '', '', '', '', '', '', '']);

  price_BYN_xlsx_array.push([
    'Картинка',
    'Модель',
    `Количество\nв оптовой\nкоробке`,
    'Розничная цена\n1 единицы\n с НДС\n в Беларуси\n с доставкой\nв белорусских\nрублях (BYN)',
    'Оптовая цена\n1 eдиницы\n с НДС\n в Беларуси\n с доставкой\nв белорусских\nрублях (BYN)',
    'Наименование',
    'Вес\nоптовой\nкоробки',
    'Объем\nоптовой\nкоробки',
    'Диаметр\n(трубы)\nподводки\nводы',
    'Гарантия',
  ]);
  price_USD_xlsx_array.push([
    'Картинка',
    'Модель',
    `Количество\nв оптовой\nкоробке`,
    'Розничная цена\n1 единицы\n с НДС\n в Стамбуле\nв долларах\nСША (USD)',
    'Оптовая цена\n1 eдиницы\n с НДС\n в Стамбуле\nв долларах\nСША (USD)',
    'Наименование',
    'Вес\nоптовой\nкоробки',
    'Объем\nоптовой\nкоробки',
    'Диаметр\n(трубы)\nподводки\nводы',
    'Гарантия',
  ]);
  price_RUB_xlsx_array.push([
    'Картинка',
    'Модель',
    `Количество\nв оптовой\nкоробке`,
    'Розничная цена\n1 единицы\n с НДС\n в России\n с доставкой\nв российских\nрублях (RUB)',
    'Оптовая цена\n1 eдиницы\n с НДС\n в России\n с доставкой\nв российских\nрублях (RUB)',
    'Наименование',
    'Вес\nоптовой\nкоробки',
    'Объем\nоптовой\nкоробки',
    'Диаметр\n(трубы)\nподводки\nводы',
    'Гарантия',
  ]);
  price_GEL_xlsx_array.push([
    'Картинка',
    'Модель',
    `Количество\nв оптовой\nкоробке`,
    'Розничная цена\n1 единицы\n с НДС\n в Грузии\n с доставкой\nв грузинских\nлари (GEL)',
    'Оптовая цена\n1 eдиницы\n с НДС\n в Грузии\n с доставкой\nв грузинских\nлари (GEL)',
    'Наименование',
    'Вес\nоптовой\nкоробки',
    'Объем\nоптовой\nкоробки',
    'Диаметр\n(трубы)\nподводки\nводы',
    'Гарантия',
  ]);
  price_AMD_xlsx_array.push([
    'Картинка',
    'Модель',
    `Количество\nв оптовой\nкоробке`,
    'Розничная цена\n1 единицы\n с НДС\n в Армении\n с доставкой\nв армянянских\nдрам (AMD)',
    'Оптовая цена\n1 eдиницы\n с НДС\n в Армении\n с доставкой\nв армянянских\nдрам (AMD)',
    'Наименование',
    'Вес\nоптовой\nкоробки',
    'Объем\nоптовой\nкоробки',
    'Диаметр\n(трубы)\nподводки\nводы',
    'Гарантия',
  ]);

  const brands = FetchItemBrands.get();

  const items = FetchItems.get();

  const categories = FetchItemCategories.get()
    .filter(e => !e.dp_isHidden)
    .sort((a, b) => a.dp_sortingIndex - b.dp_sortingIndex);

  let rowId = 2;
  const itemRowIds: number[] = [];
  const itemBrandRowIds: number[] = [];
  const itemCategoryRowIds: number[] = [];
  AppEnv.price__brands.split(',').forEach(currentEnvBrand => {
    brands.forEach(currentBrand => {
      if (currentEnvBrand === currentBrand.dp_urlSegment) {
        const level1 = `~ ~ ~ ~ ~ ~ ~ ~ ${currentBrand.dp_name} ~ ~ ~ ~ ~ ~ ~ ~ ~`;
        price_BYN_xlsx_array.push([level1, '', '', '', '', '', '', '', '', '']);
        price_USD_xlsx_array.push([level1, '', '', '', '', '', '', '', '', '']);
        price_RUB_xlsx_array.push([level1, '', '', '', '', '', '', '', '', '']);
        price_GEL_xlsx_array.push([level1, '', '', '', '', '', '', '', '', '']);
        price_AMD_xlsx_array.push([level1, '', '', '', '', '', '', '', '', '']);

        rowId += 1;
        itemBrandRowIds.push(rowId);

        categories.forEach(currentCategory => {
          if (currentBrand.dp_id === currentCategory.dp_itemBrandId) {
            const level2 = `~ ~ ~ ${currentCategory.dp_name} ~ ~ ~`;
            price_BYN_xlsx_array.push([level2, '', '', '', '', '', '', '', '', '']);
            price_USD_xlsx_array.push([level2, '', '', '', '', '', '', '', '', '']);
            price_RUB_xlsx_array.push([level2, '', '', '', '', '', '', '', '', '']);
            price_GEL_xlsx_array.push([level2, '', '', '', '', '', '', '', '', '']);
            price_AMD_xlsx_array.push([level2, '', '', '', '', '', '', '', '', '']);

            rowId += 1;
            itemCategoryRowIds.push(rowId);

            items.forEach(currentItem => {
              if (currentCategory.dp_id === currentItem.dp_itemCategoryId && !currentItem.dp_isHidden) {
                const img =
                  currentItem.dp_photoUrl.length === 0
                    ? 'нет\nкартинки'
                    : `=IMAGE("${currentItem.dp_photoUrl}")`;

                const model = currentItem.dp_model;

                const onBox = ItemObject.getOnBox(currentItem);

                const wholesaleСostostBYN = ItemObject.getWholesaleCostBYN(
                  currentItem,
                ).replace('.', ',');

                const wholesaleСostUSD = ItemObject.getWholesaleCostUSD(
                  currentItem,
                ).replace('.', ',');

                const wholesaleСostRUB = ItemObject.getWholesaleCostRUB(
                  currentItem,
                ).replace('.', ',');

                const retailСostBYN = ItemObject.getRetailCostBYN(
                  currentItem,
                ).replace('.', ',');

                const retailСostUSD = ItemObject.getRetailCostUSD(
                  currentItem,
                ).replace('.', ',');

                const retailСostRUB = ItemObject.getRetailCostRUB(
                  currentItem,
                ).replace('.', ',');

                // < < < name
                const nameMain = currentItem.dp_name;
                const nameRu = ItemObject.getRuName(currentItem);
                const nameEn = ItemObject.getEnName(currentItem);
                const nameTr = ItemObject.getTrName(currentItem);

                let name = '';

                if (nameTr.length > 0) {
                  name += `TR: ${nameTr}`;
                  name += '\n';
                }

                if (nameEn.length > 0) {
                  name += `EN: ${nameEn}`;
                  name += '\n';
                }

                if (nameRu.length > 0) {
                  name += `RU: ${nameRu}`;
                  name += '\n';
                }

                if (name.length === 0) {
                  name = nameMain;
                } else if (nameRu.length === 0) {
                  name = `${name}Own: ${nameMain}`;
                }

                name = name.trim();
                // > > > end name

                const kg = ItemObject.getKg(currentItem);
                const V = ItemObject.getV(currentItem);
                const diametr = ItemObject.getDiametr(currentItem);
                const warranty = ItemObject.getWarranty(currentItem);

                price_BYN_xlsx_array.push([
                  img,
                  model,
                  onBox.length > 0 ? onBox : 'уточняйте',
                  retailСostBYN.length > 0 ? retailСostBYN : 'уточняйте',
                  wholesaleСostostBYN.length > 0
                    ? wholesaleСostostBYN
                    : 'уточняйте',
                  name,
                  kg,
                  V,
                  diametr,
                  warranty,
                ]);
                price_USD_xlsx_array.push([
                  img,
                  model,
                  onBox.length > 0 ? onBox : 'уточняйте',
                  retailСostUSD.length > 0 ? retailСostUSD : 'уточняйте',
                  wholesaleСostUSD.length > 0 ? wholesaleСostUSD : 'уточняйте',
                  name,
                  kg,
                  V,
                  diametr,
                  warranty,
                ]);
                price_RUB_xlsx_array.push([
                  img,
                  model,
                  onBox.length > 0 ? onBox : 'уточняйте',
                  retailСostRUB.length > 0 ? retailСostRUB : 'уточняйте',
                  wholesaleСostRUB.length > 0 ? wholesaleСostRUB : 'уточняйте',
                  name,
                  kg,
                  V,
                  diametr,
                  warranty,
                ]);
                price_GEL_xlsx_array.push([
                  img,
                  model,
                  onBox.length > 0 ? onBox : 'уточняйте',
                  'скоро',
                  'скоро',
                  name,
                  kg,
                  V,
                  diametr,
                  warranty,
                ]);
                price_AMD_xlsx_array.push([
                  img,
                  model,
                  onBox.length > 0 ? onBox : 'уточняйте',
                  'скоро',
                  'скоро',
                  name,
                  kg,
                  V,
                  diametr,
                  warranty,
                ]);

                rowId += 1;
                itemRowIds.push(rowId);
              }
            });
          }
        });
      }
    });
  });

  // < < < Insert Data to Google Sheets
  const rangeBYN = listBYN.getRange(
    1,
    1,
    price_BYN_xlsx_array.length,
    price_BYN_xlsx_array[0].length,
  );
  const rangeUSD = listUSD.getRange(
    1,
    1,
    price_USD_xlsx_array.length,
    price_USD_xlsx_array[0].length,
  );
  const rangeRUB = listRUB.getRange(
    1,
    1,
    price_RUB_xlsx_array.length,
    price_RUB_xlsx_array[0].length,
  );
  const rangeGEL = listGEL.getRange(
    1,
    1,
    price_GEL_xlsx_array.length,
    price_GEL_xlsx_array[0].length,
  );
  const rangeAMD = listAMD.getRange(
    1,
    1,
    price_AMD_xlsx_array.length,
    price_AMD_xlsx_array[0].length,
  );

  rangeBYN.setValues(price_BYN_xlsx_array);
  rangeUSD.setValues(price_USD_xlsx_array);
  rangeRUB.setValues(price_RUB_xlsx_array);
  rangeGEL.setValues(price_GEL_xlsx_array);
  rangeAMD.setValues(price_AMD_xlsx_array);
  // > > > end Insert Data to Google Sheets

  // < < < set borders
  Logger.log('set borders');

  const allTableRangeBYN = listBYN.getRange(
    `A1:J${price_BYN_xlsx_array.length}`,
  );
  const allTableRangeUSD = listUSD.getRange(
    `A1:J${price_USD_xlsx_array.length}`,
  );
  const allTableRangeRUB = listRUB.getRange(
    `A1:J${price_RUB_xlsx_array.length}`,
  );
  const allTableRangeGEL = listGEL.getRange(
    `A1:J${price_GEL_xlsx_array.length}`,
  );
  const allTableRangeAMD = listAMD.getRange(
    `A1:J${price_AMD_xlsx_array.length}`,
  );

  // Аргументы setBorder(): (top, left, bottom, right, vertical, horizontal, color)
  [
    allTableRangeBYN,
    allTableRangeUSD,
    allTableRangeRUB,
    allTableRangeGEL,
    allTableRangeAMD,
  ].forEach(allTableRange => {
    allTableRange.setBorder(true, true, true, true, true, true);
  });
  // > > > end set borders

  // < < < set col width
  Logger.log('Set column width');

  [listBYN, listUSD, listRUB, listGEL, listAMD].forEach(list => {
    list.setColumnWidth(1, 100);
    list.setColumnWidth(2, 120);
    list.setColumnWidth(3, 80);
    list.setColumnWidth(4, 100);
    list.setColumnWidth(5, 100);
    list.setColumnWidth(6, 700);
    list.setColumnWidth(7, 60);
    list.setColumnWidth(8, 60);
    list.setColumnWidth(9, 60);
    list.setColumnWidth(10, 60);
  });
  // > > > end set col width

  // < < < brand styles
  Logger.log('Set item brand styles');
  for (let i = 0; i < itemBrandRowIds.length; ++i) {
    const rowNumber = itemBrandRowIds[i];

    const rangeBYN = listBYN.getRange(`A${rowNumber}:J${rowNumber}`);
    const rangeUSD = listUSD.getRange(`A${rowNumber}:J${rowNumber}`);
    const rangeRUB = listRUB.getRange(`A${rowNumber}:J${rowNumber}`);
    const rangeGEL = listGEL.getRange(`A${rowNumber}:J${rowNumber}`);
    const rangeAMD = listAMD.getRange(`A${rowNumber}:J${rowNumber}`);

    // Установка высоты строки
    const rowHeight = 20;
    [listBYN, listUSD, listRUB, listGEL, listAMD].forEach(list => {
      list.setRowHeights(rowNumber, 1, rowHeight);
    });

    // Установка выравнивания
    [rangeBYN, rangeUSD, rangeRUB, rangeGEL, rangeAMD].forEach(range => {
      range.setHorizontalAlignment('center');
      range.setVerticalAlignment('middle');
    });

    // Объединение строк
    [rangeBYN, rangeUSD, rangeRUB, rangeGEL, rangeAMD].forEach(range => {
      range.merge();
    });

    // Закрашивание ячейки
    const color = '#b6d7a8';
    [rangeBYN, rangeUSD, rangeRUB, rangeGEL, rangeAMD].forEach(range => {
      range.setBackground(color);
    });
  }
  // > > > end brand styles

  // < < < category styles
  Logger.log('Set item category styles');
  for (let i = 0; i < itemCategoryRowIds.length; ++i) {
    const rowNumber = itemCategoryRowIds[i];

    const rangeBYN = listBYN.getRange(`A${rowNumber}:J${rowNumber}`);
    const rangeUSD = listUSD.getRange(`A${rowNumber}:J${rowNumber}`);
    const rangeRUB = listRUB.getRange(`A${rowNumber}:J${rowNumber}`);
    const rangeGEL = listGEL.getRange(`A${rowNumber}:J${rowNumber}`);
    const rangeAMD = listAMD.getRange(`A${rowNumber}:J${rowNumber}`);

    // Установка высоты строки
    const rowHeight = 20;
    [listBYN, listUSD, listRUB, listGEL, listAMD].forEach(list => {
      list.setRowHeights(rowNumber, 1, rowHeight);
    });

    // Установка выравнивания
    [rangeBYN, rangeUSD, rangeRUB, rangeGEL, rangeAMD].forEach(range => {
      range.setHorizontalAlignment('center');
      range.setVerticalAlignment('middle');
    });

    // Объединение строк
    [rangeBYN, rangeUSD, rangeRUB, rangeGEL, rangeAMD].forEach(range => {
      range.merge();
    });

    // Закрашивание ячейки
    const color = '#d9ead3';
    [rangeBYN, rangeUSD, rangeRUB, rangeGEL, rangeAMD].forEach(range => {
      range.setBackground(color);
    });
  }
  // > > > end category styles

  // < < < items styles
  let progresBar = '';
  for (let i = 0; i < itemRowIds.length; ++i) {
    const percentage = getPercentage(i + 1, itemRowIds.length);
    const newProgresBar = getProgressBar(Number(percentage));
    if (progresBar !== newProgresBar) {
      progresBar = newProgresBar;
      Logger.log(
        `Items style: ${progresBar} ${percentage}% ${i + 1}/${
          itemRowIds.length
        }`,
      );
    }

    const rowNumber = itemRowIds[i];

    // Установка высоты строки
    const rowHeight = 80;
    [listBYN, listUSD, listRUB, listGEL, listAMD].forEach(list => {
      list.setRowHeights(rowNumber, 1, rowHeight);
    });

    // Установка выравнивания для столбцов "Картинка"
    const rangeImgBYN = listBYN.getRange(`A${rowNumber}:B${rowNumber}`);
    const rangeImgUSD = listUSD.getRange(`A${rowNumber}:B${rowNumber}`);
    const rangeImgRUB = listRUB.getRange(`A${rowNumber}:B${rowNumber}`);
    const rangeImgGEL = listGEL.getRange(`A${rowNumber}:B${rowNumber}`);
    const rangeImgAMD = listAMD.getRange(`A${rowNumber}:B${rowNumber}`);

    [rangeImgBYN, rangeImgUSD, rangeImgRUB, rangeImgGEL, rangeImgAMD].forEach(
      range => {
        range.setHorizontalAlignment('center');
        range.setVerticalAlignment('middle');
      },
    );

    // Установка выравнивания для столбцов "Модель"
    const rangeModelBYN = listBYN.getRange(`B${rowNumber}`);
    const rangeModelUSD = listUSD.getRange(`B${rowNumber}`);
    const rangeModelRUB = listRUB.getRange(`B${rowNumber}`);
    const rangeModelGEL = listGEL.getRange(`B${rowNumber}`);
    const rangeModelAMD = listAMD.getRange(`B${rowNumber}`);

    [
      rangeModelBYN,
      rangeModelUSD,
      rangeModelRUB,
      rangeModelGEL,
      rangeModelAMD,
    ].forEach(range => {
      range.setHorizontalAlignment('left');
      range.setVerticalAlignment('middle');
    });

    // Установка выравнивания для столбцов "В коробке"
    const rangeOnBoxBYN = listBYN.getRange(`C${rowNumber}`);
    const rangeOnBoxUSD = listUSD.getRange(`C${rowNumber}`);
    const rangeOnBoxRUB = listRUB.getRange(`C${rowNumber}`);
    const rangeOnBoxGEL = listGEL.getRange(`C${rowNumber}`);
    const rangeOnBoxAMD = listAMD.getRange(`C${rowNumber}`);

    [
      rangeOnBoxBYN,
      rangeOnBoxUSD,
      rangeOnBoxRUB,
      rangeOnBoxGEL,
      rangeOnBoxAMD,
    ].forEach(range => {
      range.setHorizontalAlignment('left');
      range.setVerticalAlignment('middle');
    });

    // Установка выравнивания для столбца "... цена ..."
    const rangeCostBYN = listBYN.getRange(`D${rowNumber}:E${rowNumber}`);
    const rangeCostUSD = listUSD.getRange(`D${rowNumber}:E${rowNumber}`);
    const rangeCostRUB = listRUB.getRange(`D${rowNumber}:E${rowNumber}`);
    const rangeCostGEL = listGEL.getRange(`D${rowNumber}:E${rowNumber}`);
    const rangeCostAMD = listAMD.getRange(`D${rowNumber}:E${rowNumber}`);

    [
      rangeCostBYN,
      rangeCostUSD,
      rangeCostRUB,
      rangeCostGEL,
      rangeCostAMD,
    ].forEach(range => {
      range.setHorizontalAlignment('right');
      range.setVerticalAlignment('middle');
    });

    // Установка выравнивания для столбца "Наименование"
    const rangeNameBYN = listBYN.getRange(`F${rowNumber}`);
    const rangeNameUSD = listUSD.getRange(`F${rowNumber}`);
    const rangeNameRUB = listRUB.getRange(`F${rowNumber}`);
    const rangeNameGEL = listGEL.getRange(`F${rowNumber}`);
    const rangeNameAMD = listAMD.getRange(`F${rowNumber}`);

    [
      rangeNameBYN,
      rangeNameUSD,
      rangeNameRUB,
      rangeNameGEL,
      rangeNameAMD,
    ].forEach(range => {
      range.setHorizontalAlignment('left');
      range.setVerticalAlignment('middle');
      range.setWrap(true);
    });

     // Установка выравнивания для столбца "... цена ..."
    const rangeCharacteristicsBYN = listBYN.getRange(`G${rowNumber}:J${rowNumber}`);
    const rangeCharacteristicsUSD = listUSD.getRange(`G${rowNumber}:J${rowNumber}`);
    const rangeCharacteristicsRUB = listRUB.getRange(`G${rowNumber}:J${rowNumber}`);
    const rangeCharacteristicsGEL = listGEL.getRange(`G${rowNumber}:J${rowNumber}`);
    const rangeCharacteristicsAMD = listAMD.getRange(`G${rowNumber}:J${rowNumber}`);

    [
      rangeCharacteristicsBYN,
      rangeCharacteristicsUSD,
      rangeCharacteristicsRUB,
      rangeCharacteristicsGEL,
      rangeCharacteristicsAMD,
    ].forEach(range => {
      range.setHorizontalAlignment('right');
      range.setVerticalAlignment('middle');
    });
  }
  // > > > end items styles

  // < < < head styles
  const rangeH1BYN = listBYN.getRange('A1:J1');
  const rangeH1USD = listUSD.getRange('A1:J1');
  const rangeH1RUB = listRUB.getRange('A1:J1');
  const rangeH1GEL = listGEL.getRange('A1:J1');
  const rangeH1AMD = listAMD.getRange('A1:J1');

  const rangeH2BYN = listBYN.getRange('A2:J2');
  const rangeH2USD = listUSD.getRange('A2:J2');
  const rangeH2RUB = listRUB.getRange('A2:J2');
  const rangeH2GEL = listGEL.getRange('A2:J2');
  const rangeH2AMD = listAMD.getRange('A2:J2');

  // Выравнивание по центру
  Logger.log('Head style: set aligment');

  [
    rangeH1BYN,
    rangeH1USD,
    rangeH1RUB,
    rangeH1GEL,
    rangeH1AMD,
    rangeH2BYN,
    rangeH2USD,
    rangeH2RUB,
    rangeH2GEL,
    rangeH2AMD,
  ].forEach(range => {
    range.setHorizontalAlignment('center');
    range.setVerticalAlignment('middle');
    range.setBackground('#f9cb9c');
  });

  // Установка высоты
  Logger.log('Head style: set row height');

  [listBYN, listUSD, listRUB, listGEL, listAMD].forEach(list => {
    list.setRowHeights(1, 1, 30);
  });

  [listBYN, listUSD, listRUB, listGEL, listAMD].forEach(list => {
    list.setRowHeights(2, 1, 120);
  });

  // Закрепляем строку
  Logger.log('Head style: set froze row');
  [listBYN, listUSD, listRUB, listGEL, listAMD].forEach(list => {
    list.setFrozenRows(2);
  });

  [rangeH1BYN, rangeH1USD, rangeH1RUB, rangeH1GEL, rangeH1AMD].forEach(
    range => {
      range.merge();
    },
  );
  // > > > end head styles

  // < < < delete null column
  [listBYN, listUSD, listRUB, listGEL, listAMD].forEach(list => {
    try {
      list.deleteColumns(11, 21);
    } catch (exception) {
      Logger.log('Not all colums deleted (colF to colZ)');
    }
  });
  // > > > end delete null column
}
