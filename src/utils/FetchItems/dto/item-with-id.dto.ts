interface Item_ItemCharacteristicsWithId {
  dp_id: number;
  dp_itemId: string;
  dp_characteristicId: number;
  dp_value: string;
}

interface Item_ItemGaleryWithId {
  dp_id: number;
  dp_itemId: string;
  dp_photoUrl: string;
}

interface ItemWithIdDto extends ItemDto {
  dp_id: string;
  dp_itemCharacteristics: Item_ItemCharacteristicsWithId[];
  dp_itemGalery: Item_ItemGaleryWithId[];
}
