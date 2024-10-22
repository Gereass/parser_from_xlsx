const xlsx = require('xlsx');
const puppeteer = require('puppeteer');

function main() {
  try {
    // Загрузка xlsx-файла
    const workbook = xlsx.readFile('data.xlsx');
    const sheet = workbook.Sheets['Sheet1'];
    const data = xlsx.utils.sheet_to_json(sheet);

    // Создание браузера и страницы
    const browser = puppeteer.launch({ headless: false, slowMo: 100 });
    const page = browser.newPage();

    // Переход на сайт
    page.goto('https://popdecor.ru/couturierweb?utm_source=op'); // замените на нужный URL

    // Цикл по записям
    for (const row of data) {
      // Заполнение полей и отправка формы
      page.type('input[placeholder="Номер телефона с кодом страны*"]', row.phone);
      page.type('input[placeholder="E-mail*"]', row.email);
      page.click('.btn-new.btn-submit');
      page.goBack(); // Вернуться на страницу назад

      // Задержка в 2 секунды
      setTimeout(() => {}, 2000);
    }

    // Закрытие браузера
    browser.close();
  } catch (error) {
    console.error(error);
  }
}

main();
