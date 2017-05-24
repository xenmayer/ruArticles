# Парсим xls и возвращаем результат с помощью PHPExcel (http://www.codeplex.com/PHPExcel)

Задача:
Есть файл xls, в котором есть ячейки с данными формата ${key}, которые надо заменить на нужные нам.

| First Header  | Second Header |
| ------------- | ------------- |
| Content Cell  | ${key1}  |
| Content Cell  | ${key2}  |
Предлагаю один из способов решения:


```php
    /**
     * Массив исходных данных, которые должны вставляться вместо ${keyN},
     * причем ключи массива должны соответствовать keyN
     */
    $fields = [
        'key1' => 99, 
        'key2' => 222, 
    ];

    /**
     * Создаем экземпляр класса PHPExcel_IOFactory и передаем ему путь до нашего шаблонного xls
     */
    $template_file_path = "/path/to/template_file.xls";
    $php_excel = PHPExcel_IOFactory::load($template_file_path);

    /**
     * Устанавливаем активный лист
     */
    $php_excel->setActiveSheetIndex(0);

    /**
     * Получаем активный лист
     */   
    $active_sheet = $php_excel->getActiveSheet();

    /**
     * Циклом пробегаемся по всем ячейкам
     */
    foreach($active_sheet->getRowIterator() as $row) {

        foreach($row->getCellIterator() as $cell) {

            /**
             * С помощью регулярного выражения находим ключ
             */
            $key = preg_match('/\${(.*)}/', $cell->getValue(), $matches);

            /**
             * Если элемент с таким ключем существует и не пуст, то пишем его в ячейку
             * если пуст или отсутствует - пишем пустую ячейку
             */
            if ($fields[$matches[1]]) {
                $cell->setValue($fields[$matches[1]]);
            } elseif ($key) {
                $cell->setValue('');
            }

        }

    }

    /**
     * Имя возвращаемого файла
     */   
    $output_filename = 'filename.xls';

    header('Content-type: application/vnd.ms-excel');
    header('Content-Disposition: attachment; filename="' . $output_filename . '"');

    $php_excel_writer = PHPExcel_IOFactory::createWriter($php_excel, 'Excel5');
    $php_excel_writer->save('php://output');
```
