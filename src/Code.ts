const AppEnv = {
  price__google_sheet_id: '1ksxZm9z4Ix4H6O1Z-m4mqxGdms3N3aGvp46jM_jFve4',
  price__brands: ['mega', 'mega-general'],
};

function myFunction() {
  const table = SpreadsheetApp.openById(AppEnv.price__google_sheet_id);
  const listBYN = table.getSheetByName('BYN');
  const listUSD = table.getSheetByName('USD');
  const listRUB = table.getSheetByName('RUB');
  listBYN.clear();
  listUSD.clear();
  listRUB.clear();

  const price_BYN_xlsx_array: string[][] = [];
  const price_USD_xlsx_array: string[][] = [];
  const price_RUB_xlsx_array: string[][] = [];

  price_BYN_xlsx_array.push([
    'Картинка',
    'Модель',
    `Количество\nв\nящике`,
    'Цена 1 шт.\n с НДС\n в Беларуси\n с доставкой',
    'Наименование',
  ]);
  price_USD_xlsx_array.push([
    'Картинка',
    'Модель',
    `Количество\nв\nящике`,
    'Цена 1 шт.\n в Стамбуле',
    'Наименование',
  ]);
  price_RUB_xlsx_array.push([
    'Картинка',
    'Модель',
    `Количество\nв\nящике`,
    'Цена 1 шт.\n с НДС\n в РФ\n с доставкой',
    'Наименование',
  ]);

  const brands = FetchItemBrands.get();

  const items = FetchItems.get();

  const categories = FetchItemCategories.get()
    .filter(e => !e.dp_isHidden)
    .sort((a, b) => a.dp_sortingIndex - b.dp_sortingIndex);

  let rowId = 1;
  const itemRowIds: number[] = [];
  const itemBrandRowIds: number[] = [];
  const itemCategoryRowIds: number[] = [];
  AppEnv.price__brands.forEach(currentEnvBrand => {
    brands.forEach(currentBrand => {
      if (currentEnvBrand === currentBrand.dp_urlSegment) {
        const level1 = `~ ~ ~ ~ ~ ~ ~ ~ ${currentBrand.dp_name} ~ ~ ~ ~ ~ ~ ~ ~ ~`;
        price_BYN_xlsx_array.push([level1, '', '', '', '']);
        price_USD_xlsx_array.push([level1, '', '', '', '']);
        price_RUB_xlsx_array.push([level1, '', '', '', '']);

        rowId += 1;
        itemBrandRowIds.push(rowId);

        categories.forEach(currentCategory => {
          if (currentBrand.dp_id === currentCategory.dp_itemBrandId) {
            const level2 = `~ ~ ~ ${currentCategory.dp_name} ~ ~ ~`;
            price_BYN_xlsx_array.push([level2, '', '', '', '']);
            price_USD_xlsx_array.push([level2, '', '', '', '']);
            price_RUB_xlsx_array.push([level2, '', '', '', '']);

            rowId += 1;
            itemCategoryRowIds.push(rowId);

            items.forEach(currentItem => {
              if (currentCategory.dp_id === currentItem.dp_itemCategoryId) {
                const img =
                  currentItem.dp_photoUrl.length === 0
                    ? 'нет\nкартинки'
                    : `=IMAGE("${currentItem.dp_photoUrl}")`;
                const model = currentItem.dp_model;

                // < < < onBox
                const onBox = ItemObject.getOnBox(currentItem);
                const resultOnBox = onBox.length > 0 ? onBox : 'уточняйте';
                // > > > end onBox

                // < < < cost
                // const cost = Number(currentItem.dp_cost).toFixed(2);
                const costBYN = ItemObject.getCostBYN(currentItem);
                const costUSD = ItemObject.getCostUSD(currentItem);
                const costRUB = ItemObject.getCostRUB(currentItem);
                const resultBYN = costBYN.length > 0 ? costBYN : 'уточняйте';
                const resultUSD = costUSD.length > 0 ? costUSD : 'уточняйте';
                const resultRUB = costRUB.length > 0 ? costRUB : 'уточняйте';
                // > > > end cost

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
                }

                name = name.trim();
                // > > > end name

                price_BYN_xlsx_array.push([
                  img,
                  model,
                  resultOnBox,
                  resultBYN,
                  name,
                ]);
                price_USD_xlsx_array.push([
                  img,
                  model,
                  resultOnBox,
                  resultUSD,
                  name,
                ]);
                price_RUB_xlsx_array.push([
                  img,
                  model,
                  resultOnBox,
                  resultRUB,
                  name,
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

  rangeBYN.setValues(price_BYN_xlsx_array);
  rangeUSD.setValues(price_USD_xlsx_array);
  rangeRUB.setValues(price_RUB_xlsx_array);
  // > > > end Insert Data to Google Sheets

  // < < < set borders
  Logger.log('set borders');

  const allTableRangeBYN = listBYN.getRange(
    `A1:E${price_BYN_xlsx_array.length}`,
  );
  const allTableRangeUSD = listUSD.getRange(
    `A1:E${price_USD_xlsx_array.length}`,
  );
  const allTableRangeRUB = listRUB.getRange(
    `A1:E${price_RUB_xlsx_array.length}`,
  );

  // Аргументы setBorder(): (top, left, bottom, right, vertical, horizontal, color)
  allTableRangeBYN.setBorder(true, true, true, true, true, true);
  allTableRangeUSD.setBorder(true, true, true, true, true, true);
  allTableRangeRUB.setBorder(true, true, true, true, true, true);
  // > > > end set borders

  // < < < set col width
  Logger.log('Set column width');

  listBYN.setColumnWidth(1, 100);
  listUSD.setColumnWidth(1, 100);
  listRUB.setColumnWidth(1, 100);

  listBYN.setColumnWidth(2, 100);
  listUSD.setColumnWidth(2, 100);
  listRUB.setColumnWidth(2, 100);

  listBYN.setColumnWidth(3, 100);
  listUSD.setColumnWidth(3, 100);
  listRUB.setColumnWidth(3, 100);

  listBYN.setColumnWidth(4, 100);
  listUSD.setColumnWidth(4, 100);
  listRUB.setColumnWidth(4, 100);

  listBYN.setColumnWidth(5, 700);
  listUSD.setColumnWidth(5, 700);
  listRUB.setColumnWidth(5, 700);
  // > > > end set col width

  // < < < brand styles
  Logger.log('Set item brand styles');
  for (let i = 0; i < itemBrandRowIds.length; ++i) {
    const rowNumber = itemBrandRowIds[i];

    const rangeBYN = listBYN.getRange(`A${rowNumber}:E${rowNumber}`);
    const rangeUSD = listUSD.getRange(`A${rowNumber}:E${rowNumber}`);
    const rangeRUB = listRUB.getRange(`A${rowNumber}:E${rowNumber}`);

    // Установка высоты строки
    const rowHeight = 20;
    listBYN.setRowHeights(rowNumber, 1, rowHeight);
    listUSD.setRowHeights(rowNumber, 1, rowHeight);
    listRUB.setRowHeights(rowNumber, 1, rowHeight);

    // Установка выравнивания
    rangeBYN.setHorizontalAlignment('center');
    rangeUSD.setHorizontalAlignment('center');
    rangeRUB.setHorizontalAlignment('center');

    rangeBYN.setVerticalAlignment('middle');
    rangeUSD.setVerticalAlignment('middle');
    rangeRUB.setVerticalAlignment('middle');

    // Объединение строк
    rangeBYN.merge();
    rangeUSD.merge();
    rangeRUB.merge();

    // Закрашивание ячейки
    const color = '#b6d7a8';
    rangeBYN.setBackground(color);
    rangeUSD.setBackground(color);
    rangeRUB.setBackground(color);
  }
  // > > > end brand styles

  // < < < category styles
  Logger.log('Set item category styles');
  for (let i = 0; i < itemCategoryRowIds.length; ++i) {
    const rowNumber = itemCategoryRowIds[i];

    const rangeBYN = listBYN.getRange(`A${rowNumber}:E${rowNumber}`);
    const rangeUSD = listUSD.getRange(`A${rowNumber}:E${rowNumber}`);
    const rangeRUB = listRUB.getRange(`A${rowNumber}:E${rowNumber}`);

    // Установка высоты строки
    const rowHeight = 20;
    listBYN.setRowHeights(rowNumber, 1, rowHeight);
    listUSD.setRowHeights(rowNumber, 1, rowHeight);
    listRUB.setRowHeights(rowNumber, 1, rowHeight);

    // Установка выравнивания
    rangeBYN.setHorizontalAlignment('center');
    rangeUSD.setHorizontalAlignment('center');
    rangeRUB.setHorizontalAlignment('center');

    rangeBYN.setVerticalAlignment('middle');
    rangeUSD.setVerticalAlignment('middle');
    rangeRUB.setVerticalAlignment('middle');

    // Объединение строк
    rangeBYN.merge();
    rangeUSD.merge();
    rangeRUB.merge();

    // Закрашивание ячейки
    const color = '#d9ead3';
    rangeBYN.setBackground(color);
    rangeUSD.setBackground(color);
    rangeRUB.setBackground(color);
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
    const rowHeight = 60;
    listBYN.setRowHeights(rowNumber, 1, rowHeight);
    listUSD.setRowHeights(rowNumber, 1, rowHeight);
    listRUB.setRowHeights(rowNumber, 1, rowHeight);

    // Установка выравнивания для столбцов "Картинка", "Наименование"
    const rangeImgAndModelBYN = listBYN.getRange(`A${rowNumber}:B${rowNumber}`);
    const rangeImgAndModelUSD = listUSD.getRange(`A${rowNumber}:B${rowNumber}`);
    const rangeImgAndModelRUB = listRUB.getRange(`A${rowNumber}:B${rowNumber}`);

    rangeImgAndModelBYN.setHorizontalAlignment('center');
    rangeImgAndModelUSD.setHorizontalAlignment('center');
    rangeImgAndModelRUB.setHorizontalAlignment('center');

    rangeImgAndModelBYN.setVerticalAlignment('middle');
    rangeImgAndModelUSD.setVerticalAlignment('middle');
    rangeImgAndModelRUB.setVerticalAlignment('middle');

    // Установка выравнивания для столбца "Стоимость ..."
    const rangeCostBYN = listBYN.getRange(`C${rowNumber}:D${rowNumber}`);
    const rangeCostUSD = listUSD.getRange(`C${rowNumber}:D${rowNumber}`);
    const rangeCostRUB = listRUB.getRange(`C${rowNumber}:D${rowNumber}`);

    rangeCostBYN.setHorizontalAlignment('right');
    rangeCostUSD.setHorizontalAlignment('right');
    rangeCostRUB.setHorizontalAlignment('right');

    rangeCostBYN.setVerticalAlignment('middle');
    rangeCostUSD.setVerticalAlignment('middle');
    rangeCostRUB.setVerticalAlignment('middle');

    // Установка выравнивания для столбца "Наименование"
    const rangeNameBYN = listBYN.getRange(`E${rowNumber}`);
    const rangeNameUSD = listUSD.getRange(`E${rowNumber}`);
    const rangeNameRUB = listRUB.getRange(`E${rowNumber}`);

    rangeNameBYN.setHorizontalAlignment('left');
    rangeNameUSD.setHorizontalAlignment('left');
    rangeNameRUB.setHorizontalAlignment('left');

    rangeNameBYN.setVerticalAlignment('middle');
    rangeNameUSD.setVerticalAlignment('middle');
    rangeNameRUB.setVerticalAlignment('middle');

    rangeNameBYN.setWrap(true);
    rangeNameUSD.setWrap(true);
    rangeNameRUB.setWrap(true);
  }
  // > > > end items styles

  // < < < head styles
  // Выравнивание по центру
  Logger.log('Head style: set aligment');
  const headBYN = listBYN.getRange('A1:E1');
  const headUSD = listUSD.getRange('A1:E1');
  const headRUB = listRUB.getRange('A1:E1');

  headBYN.setHorizontalAlignment('center');
  headUSD.setHorizontalAlignment('center');
  headRUB.setHorizontalAlignment('center');

  headBYN.setVerticalAlignment('middle');
  headUSD.setVerticalAlignment('middle');
  headRUB.setVerticalAlignment('middle');

  // Установка высоты
  Logger.log('Head style: set row height');
  const rowNumber = 1;

  const rangeNameBYN = listBYN.getRange(`A${rowNumber}:E${rowNumber}`);
  const rangeNameUSD = listUSD.getRange(`A${rowNumber}:E${rowNumber}`);
  const rangeNameRUB = listRUB.getRange(`A${rowNumber}:E${rowNumber}`);

  const rowHeight = 80;
  listBYN.setRowHeights(rowNumber, 1, rowHeight);
  listUSD.setRowHeights(rowNumber, 1, rowHeight);
  listRUB.setRowHeights(rowNumber, 1, rowHeight);

  // Закрепляем строку
  Logger.log('Head style: set froze row');
  listBYN.setFrozenRows(1);
  listUSD.setFrozenRows(1);
  listRUB.setFrozenRows(1);

  // Закрашивание ячейки
  const color = '#f9cb9c';
  rangeNameBYN.setBackground(color);
  rangeNameUSD.setBackground(color);
  rangeNameRUB.setBackground(color);
  // > > > end head styles

  // < < < delete null column
  try {
    listBYN.deleteColumns(6, 21);
  } catch (exception) {
    Logger.log('Not all colums deleted (colE to colZ)');
  }

  try {
    listUSD.deleteColumns(6, 21);
  } catch (exception) {
    Logger.log('Not all colums deleted (colE to colZ)');
  }

  try {
    listRUB.deleteColumns(6, 21);
  } catch (exception) {
    Logger.log('Not all colums deleted (colE to colZ)');
  }
  // > > > end delete null column
}
