const xlsx = require('xlsx');
const puppeteer = require('puppeteer');
const { resolve } = require('path');

(async () => {
  // Загрузка xlsx-файла
  const workbook = xlsx.readFile('data.xlsx');
  const sheet = workbook.Sheets['Sheet1'];
  const data = xlsx.utils.sheet_to_json(sheet);

  // Создание браузера и страницы
  const browser = await puppeteer.launch();
  const page = await browser.newPage();

  // Переход на сайт
  await page.goto('https://popdecor.ru/couturierweb?utm_source=op'); // замените на нужный URL

  // Заполнение полей и отправка формы
  for (const row of data) {
    await page.type('input[placeholder="Номер телефона с кодом страны*"]', row.phone);
    await page.type('input[placeholder="E-mail*"]', row.email);
    await page.click('.btn-new.btn-submit');
    await page.goBack();
    await new Promise(resolve => setTimeout(resolve, 2000))
    await page.waitForNavigation();
  }

  // Закрытие браузера
  await browser.close();
})();
