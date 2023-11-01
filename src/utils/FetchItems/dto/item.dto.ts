interface Item_ItemCharacteristics {
  dp_characteristicId: number;
  dp_value: string;
}

interface Item_ItemGalery {
  dp_photoUrl: string;
}

interface ItemDto {
  dp_name: string;
  dp_model: string;
  dp_cost: number;
  dp_photoUrl: string;
  dp_seoKeywords: string;
  dp_seoDescription: string;
  dp_itemCategoryId: number;
  dp_isHidden: string;
  dp_itemCharacteristics: Item_ItemCharacteristics[];
  dp_itemGalery: Item_ItemGalery[];
}
