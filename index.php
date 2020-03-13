<?php


use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

require 'vendor/autoload.php';

include_once 'lib/simple_html_dom.php';
include_once 'lib/cURL_query.php';

setLocale(LC_ALL, 'ru_RU.CP1251');
date_default_timezone_set('UTC');

//set_time_limit(4500); //  в секундах время исполения скрипта

class Parser {

    public $referer = 'k-h.com.ua/';
    public $cookiefile = '/home/andrew/www/parserAutoParts/cookie.txt';
    public $fileSourcePath = '/home/andrew/www/parserAutoParts/pricelist_zenitauto_2589.xlsx';
    public $fileResultPath = '/home/andrew/www/parserAutoParts/_result.csv';
    public $parsingDate;

// формирование прайса
    public function CreateFileToDownload() {
        $fileCreate = curl_init();

        curl_setopt($fileCreate, CURLOPT_USERAGENT, 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:67.0) Gecko/20100101 Firefox/67.0');
        curl_setopt($fileCreate, CURLOPT_URL, 'http://k-h.com.ua/cmd.php?key=zaHgjyt554rr4g5hnvt65&r=mailing/mailing/ex/price24mobis@gmail.com/2589');
        curl_setopt($fileCreate, CURLOPT_REFERER, 'http://k-h.com.ua/index.php?module=profile|price');
        curl_setopt($fileCreate, CURLOPT_FOLLOWLOCATION, true);
        curl_setopt($fileCreate, CURLOPT_COOKIE, 'client_hash=71dd1b389ea124d3c1b966509d2c0837;client_id=2589;route=%D0%9A%D0%9C3');
        curl_setopt($fileCreate, CURLOPT_RETURNTRANSFER, true);

        $result = curl_exec($fileCreate);

        //Отлавливаем ошибки подключения
        if ($result === false) {
            echo "Ошибка CURL: " . curl_error($curl);
        } else {
            curl_close($fileCreate);
            return true;
        }
    }

//скачивание файла
    public function GetFile() {

        $fp = fopen($this->fileSourcePath, "w");

        $fileDownload = curl_init();

        curl_setopt($fileDownload, CURLOPT_USERAGENT, 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:67.0) Gecko/20100101 Firefox/67.0');
        curl_setopt($fileDownload, CURLOPT_URL, 'http://k-h.com.ua/public/downloads/pricelist_zenitauto_2589.xlsx');
        curl_setopt($fileDownload, CURLOPT_REFERER, 'http://k-h.com.ua/index.php?module=profile|price');
        curl_setopt($fileDownload, CURLOPT_FOLLOWLOCATION, true);
        curl_setopt($fileDownload, CURLOPT_COOKIE, 'client_hash=71dd1b389ea124d3c1b966509d2c0837;client_id=2589;route=%D0%9A%D0%9C3');
        curl_setopt($fileDownload, CURLOPT_RETURNTRANSFER, true);
        curl_setopt($fileDownload, CURLOPT_FILE, $fp); // сюда будет загружен файл

        curl_exec($fileDownload);
        curl_close($fileDownload);

        return fclose($fp);
    }

// готовим исходный файл. Отбираем нужные бренды
    public function PreparingFileForParsing() {


        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

// Читаем файл и записываем информацию в переменную
        $spreadsheet = $reader->load($this->fileSourcePath);

        $loadedSheetNames = $spreadsheet->getSheetNames();
        $writer = new Csv($spreadsheet);
        $writer->setDelimiter(',');

        foreach ($loadedSheetNames as $sheetIndex => $loadedSheetName) {
            $writer->setSheetIndex($sheetIndex);
            $writer->save('/home/andrew/www/parserAutoParts/' . $loadedSheetName . '.csv');
        }

        $handle = fopen($loadedSheetName . '.csv', "r");
        $brandSelected = fopen('/home/andrew/www/parserAutoParts/brand-selected.csv', "w");
        while (($data = fgetcsv($handle, ",")) !== FALSE) {

            if (isset($data[3]) && $data[3] == 'MOBIS' || $data[3] == 'HYUNDAI' || $data[3] == 'Renault' || $data[3] == 'Renault Turkey') {
                fputcsv($brandSelected, array($data[0], $data[2], $data[3]), ',');
            }
        }
        fclose($handle);
        fclose($brandSelected);

        return true;
    }

    public function ParsingStockBalance() {
        $parsingDate = date('l jS \of F Y h:i:s A');
// работаем с подготовленным  файлом 
        $filePreparingOpen = fopen('/home/andrew/www/parserAutoParts/brand-selected.csv', "r");
        $fileResultOpen = fopen($this->fileResultPath, "w");

        fputcsv($fileResultOpen, array($parsingDate), ',');
        fputcsv($fileResultOpen, array('Parts Number', 'Stock Ballance', 'Brand', 'StockCity'), ',');
        if (is_readable('/home/andrew/www/parserAutoParts/brand-selected.csv')) {

            while (($data = fgetcsv($filePreparingOpen, ',')) !== FALSE) {
                $partNumber = $data[1];

                if (isset($partNumber)) {
                
                    $stockBalananceRequest = curl_init();

                    $partNumber = preg_replace("/ /", '+', $partNumber);
// делаем сURl запрос для получения точного остатка по номенклатуре
                    curl_setopt($stockBalananceRequest, CURLOPT_USERAGENT, 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:67.0) Gecko/20100101 Firefox/67.0');

                    curl_setopt($stockBalananceRequest, CURLOPT_URL, "http://k-h.com.ua/index.php?module=catalog%7Csearch&s=fc&pos=$partNumber");

                    curl_setopt($stockBalananceRequest, CURLOPT_REFERER, 'http://k-h.com.ua/index.php?module=catalog|direction&s=fc');

                    curl_setopt($stockBalananceRequest, CURLOPT_FOLLOWLOCATION, true);
                    curl_setopt($stockBalananceRequest, CURLOPT_COOKIE, 'client_hash=71dd1b389ea124d3c1b966509d2c0837;client_id=2589;route=%D0%9A%D0%9C3');
                    curl_setopt($stockBalananceRequest, CURLOPT_RETURNTRANSFER, true); // парсим в переменную

                    $resultFormSearch = curl_exec($stockBalananceRequest);
                   
// парсим форму
//                    echo $partNumber;
                    $page = str_get_html($resultFormSearch);

                    $stockCityCollection = $page->find('tr');

                    foreach ($stockCityCollection as $city) {
                        if (!isset($city->class)) {

                            $stockCity = trim($city->plaintext);

                            if (preg_match('/Киев/', $stockCity)) {
                                $stockCity = 'Kyiv';
                            } elseif (preg_match('/Харьков/', $stockCity)) {
                                $stockCity = 'Kharkiv';
                            } elseif (preg_match('/Днепр/', $stockCity)) {
                                $stockCity = 'Dnipro';
                            }
                            $nextSibling = $city->next_sibling();

                            do {
// находим таблицы с классом .list_table у них все формы с методом "POST" и у 
// них инпуты. Получаем значения артикула и максимального количества на складе

                                $brand = $nextSibling->find('td', 1)->plaintext; //бренд

                                if ($brand == 'MOBIS' || $brand == 'HYUNDAI' || $brand == 'Renault' || $brand == 'Renault Turkey') {
                                    $partsNumber = $nextSibling->find('td', 0)->plaintext; //артикул
                                    $stockBallance = $nextSibling->find('td form[method=POST]', 0)->find('input[name=Cmax]', 0)->value; //остаток на складе
//добавляем строки в CSV файл с разделителем ","
                                 fputcsv($fileResultOpen, array($partsNumber, $stockBallance, $brand, $stockCity), ',');
                                }

//переход к следующему соседу
                                $nextSibling = $nextSibling->next_sibling();
                            } while ($nextSibling->class == 'list_table');
                        }
                    }
                    curl_close($stockBalananceRequest);
                }
            }
        }
        fclose($filePreparingOpen);
        fclose($fileResultOpen);
        return true;
    }

    public function SendMailData() {
        $parsingDate = date('l jS \of F Y h:i:s A');

        $mail = new PHPMailer(true);

        try {
            //Server settings
            $mail->SMTPDebug = 0;                                       // Enable verbose debug output
            $mail->isSMTP();                                            // Set mailer to use SMTP
            $mail->Host = 'smtp.gmail.com';  // Specify main and backup SMTP servers
            $mail->SMTPAuth = true;                                   // Enable SMTP authentication
            $mail->Username = '';                     // SMTP username
            $mail->Password = '';                               // SMTP password
            $mail->SMTPSecure = 'tls';                                  // Enable TLS encryption, `ssl` also accepted
            $mail->Port = 587;                                    // TCP port to connect to
            //Recipients
            $mail->setFrom('from@example.com', 'Mailer');
            $mail->addAddress('avto-sklad.price@ukr.net', 'AvtoSklad');     // Add a recipient
            $mail->addCC('andymolchanov@ukr.net', 'Admin');     // Add a recipient
            // Attachments
            $mail->addAttachment('/home/andrew/www/parserAutoParts/_result.csv');         // Add attachments
            // Content
            $mail->isHTML(true);                                  // Set email format to HTML
            $mail->Subject = 'Parsing' . $parsingDate;
            $mail->Body = 'Daily report';

            $mail->send();
        } catch (Exception $e) {
//            echo "Message could not be sent. Mailer Error: {$mail->ErrorInfo}";
        }
    }

}

$parser = new Parser;
$resultCreateFileToDownload = $parser->CreateFileToDownload();
$resultGetFile = $parser->GetFile();
$resultPreparing = $parser->PreparingFileForParsing();
$resultParsing = $parser->ParsingStockBalance();
$resultSendMail = $parser->SendMailData();
