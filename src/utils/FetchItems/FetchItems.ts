class FetchItems {
  static get() {
    const url = 'https://de-pa.by/api/v1/items';
    Logger.log(`GET ${url}`);

    try {
      const response = UrlFetchApp.fetch(url);
      const statusCode = response.getResponseCode();

      if (statusCode == 200) {
        Logger.log(`${statusCode} | GET ${url}`);
        const data = response.getContentText();
        const jsonData: GetItemDto[] = JSON.parse(data);
        return jsonData;
      }

      Logger.log(`${statusCode} | GET ${url}`);
      return [];
    } catch (e) {
      Logger.log('Произошла ошибка: ' + e.toString());
      return [];
    }
  }
}
