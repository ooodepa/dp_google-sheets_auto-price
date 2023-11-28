class ItemObject {
  static getOnBox(item: ItemDto) {
    return ItemObject.getCharacteristicById(1, item);
  }

  static getRuName(item: ItemDto) {
    return ItemObject.getCharacteristicById(20, item);
  }

  static getTrName(item: ItemDto) {
    return ItemObject.getCharacteristicById(18, item);
  }

  static getEnName(item: ItemDto) {
    return ItemObject.getCharacteristicById(19, item);
  }

  static getWholesaleCostBYN(item: ItemDto) {
    return ItemObject.getCharacteristicById(25, item);
  }

  static getWholesaleCostUSD(item: ItemDto) {
    return ItemObject.getCharacteristicById(24, item);
  }

  static getWholesaleCostRUB(item: ItemDto) {
    return ItemObject.getCharacteristicById(29, item);
  }

  static getRetailCostBYN(item: ItemDto) {
    return ItemObject.getCharacteristicById(35, item);
  }

  static getRetailCostUSD(item: ItemDto) {
    return ItemObject.getCharacteristicById(37, item);
  }

  static getRetailCostRUB(item: ItemDto) {
    return ItemObject.getCharacteristicById(36, item);
  }

  static getCharacteristicById(characteristicId: number, item: ItemDto) {
    const characteristics = item.dp_itemCharacteristics;
    for (let i = 0; i < characteristics.length; ++i) {
      const currentCh = characteristics[i];
      if (currentCh.dp_characteristicId === characteristicId) {
        return currentCh.dp_value;
      }
    }
    return '';
  }
}
