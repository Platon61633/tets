const express = require('express');
const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;

// Функция для парсинга данных
async function parsePMIData() {
  try {
    const url = 'https://cpk.msu.ru/rating/dep_02';
    const response = await axios.get(url, {
      timeout: 10000,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
      }
    });
    
    const $ = cheerio.load(response.data);
    const date = $('p').eq(15).text().trim();
    const table = $('table').eq(8);
    
    const tableText = table.text().split('\n')
      .map(item => item.trim())
      .filter(item => item !== '');
    
    const data = {
      date: date,
      headers: ['номер', 'согласие', 'приоритет', 'баллы', 'статус'],
      rows: []
    };
    
    for (let i = 16; i < tableText.length - 19; i += 19) {
      const block = tableText.slice(i, i + 19);
      data.rows.push([
        block[1],  // номер
        block[2],  // согласие
        block[3],  // приоритет
        block[7],  // баллы
        block[16]  // статус
      ]);
    }
    
    return data;
    
  } catch (error) {
    console.error('Ошибка при парсинге:', error);
    throw new Error('Не удалось получить данные с сайта МГУ');
  }
}

// Маршрут для скачивания Excel-файла
app.get('/download-pmi-excel', async (req, res) => {
  try {
    // Парсим данные
    const pmiData = await parsePMIData();
    
    // Создаем Excel-книгу в памяти
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Рейтинг ПМИ');
    
    // Добавляем заголовки
    worksheet.addRow([...pmiData.headers, 'Дата обновления']);
    
    // Добавляем данные
    pmiData.rows.forEach(row => {
      worksheet.addRow([...row, pmiData.date]);
    });
    
    // Форматируем заголовки
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD9D9D9' }
    };
    
    // Настраиваем ширину колонок
    worksheet.columns = [
      { width: 10 }, // номер
      { width: 12 }, // согласие
      { width: 12 }, // приоритет
      { width: 10 }, // баллы
      { width: 20 }, // статус
      { width: 25 }  // дата
    ];
    
    // Устанавливаем заголовки ответа для скачивания файла
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename=PMI_Rating.xlsx'
    );
    
    // Отправляем файл
    await workbook.xlsx.write(res);
    res.end();
    
  } catch (error) {
    console.error('Ошибка при создании Excel:', error);
    res.status(500).send(`
      <h1>Ошибка сервера</h1>
      <p>${error.message}</p>
      <p>Попробуйте позже или обратитесь к администратору</p>
    `);
  }
});

// Стартовая страница с инструкцией
app.get('/', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>Рейтинг ПМИ МГУ</title>
      <style>
        body { font-family: Arial, sans-serif; max-width: 800px; margin: 40px auto; }
        h1 { color: #1a3d6d; }
        .btn { 
          display: inline-block; 
          padding: 15px 30px; 
          background: #1a3d6d; 
          color: white; 
          text-decoration: none; 
          border-radius: 5px;
          font-size: 18px;
          margin: 20px 0;
        }
        .btn:hover { background: #0d2a4d; }
      </style>
    </head>
    <body>
      <h1>Рейтинг поступающих на ПМИ МГУ</h1>
      <p>Сервер автоматически парсит данные с официального сайта МГУ и формирует Excel-файл</p>
      <a href="/download-pmi-excel" class="btn">Скачать Excel-файл</a>
      <p><small>При проблемах со скачиванием обновите страницу или попробуйте позже</small></p>
    </body>
    </html>
  `);
});

// Запуск сервера
app.listen(PORT, () => {
  console.log(`Сервер запущен на порту ${PORT}`);
  console.log(`Откройте в браузере: http://localhost:${PORT}`);
  console.log(`Ссылка для скачивания: http://localhost:${PORT}/download-pmi-excel`);
});