class DateController {
  static getMyFullDate(data: Date = new Date()) {
    const YYYY = data.getFullYear();
    const MM = String(data.getMonth() + 1).padStart(2, '0');
    const DD = String(data.getDate()).padStart(2, '0');

    const hh = String(data.getHours()).padStart(2, '0');
    const mm = String(data.getMinutes()).padStart(2, '0');

    return `${YYYY}-${MM}-${DD} ${hh}:${mm}`;
  }
}
