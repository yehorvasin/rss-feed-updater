function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('RSS Import')
      .addItem('Update', 'updateSheetWithRSS')
      .addToUi();
}

function updateSheetWithRSS() {
  const url = 'https://rss.app/feeds/gFz6xLcOt7CYwno8.xml';
  const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const responseCode = response.getResponseCode();

  if (responseCode === 200) {
    const xml = XmlService.parse(response.getContentText());
    const root = xml.getRootElement();
    const channel = root.getChild('channel');

    if (channel) {
      const items = channel.getChildren('item');

      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      sheet.clear();

      // Set headers
      const headers = ['Title', 'Description', 'Link', 'Publication Date'];
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f4cccc');
      headerRange.setFontSize(12);
      headerRange.setHorizontalAlignment('center');
      sheet.setRowHeight(1, 30); // Висота рядка для заголовків

      // Set data
      const data = items.map(item => [
        item.getChild('title') ? item.getChild('title').getText() : 'No title',
        item.getChild('description') ? item.getChild('description').getText() : 'No description',
        item.getChild('link') ? item.getChild('link').getText() : 'No link',
        item.getChild('pubDate') ? item.getChild('pubDate').getText() : 'No date'
      ]);

const dataRange = sheet.getRange(2, 1, data.length, headers.length);
      dataRange.setValues(data);
      dataRange.setFontSize(10);
      dataRange.setVerticalAlignment('top');
      dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Дозволити перенесення тексту

      for (let i = 2; i <= data.length + 1; i++) {
        sheet.setRowHeight(i, 50); // Висота рядка для даних
      }

      // Set column widths
      const columnWidths = [150, 300, 100, 150];
      for (let i = 0; i < headers.length; i++) {
        sheet.setColumnWidth(i + 1, columnWidths[i]);
      }

      // Set link format and alignment
      for (let i = 2; i <= data.length + 1; i++) {
        const linkCell = sheet.getRange(i, 3);
        const link = linkCell.getValue();
        linkCell.setFormula('=HYPERLINK("' + link + '", "Link")');
        linkCell.setHorizontalAlignment('center');
      }

      // Add padding by increasing column widths
      const range = sheet.getRange(2, 1, data.length, headers.length);
      range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      range.setVerticalAlignment('top');
    } else {
      SpreadsheetApp.getUi().alert('Failed to parse RSS feed. Channel element not found.');
    }
  } else {
    SpreadsheetApp.getUi().alert('Failed to fetch RSS feed. Response code: ' + responseCode);
  }
}