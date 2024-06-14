<?php
if ($_SERVER["REQUEST_METHOD"] == "POST") {
    $churchName = $_POST['churchName'];
    $reporter = $_POST['reporter'];
    $personName = $_POST['personName'];
    $gender = $_POST['gender'];
    $ageGroup = $_POST['ageGroup'];
    $status = $_POST['status'];
    $address = $_POST['address'];
    $remarks = $_POST['remarks'];

    // メール送信
    $to = 'example@example.com'; // 送信先メールアドレスを指定
    $subject = 'にをいがけ報告フォームの送信';
    $message = "教会名: $churchName\n報告者: $reporter\n相手のお名前: $personName\n性別: $gender\n年代: $ageGroup\n身上事情: $status\nご住所: $address\n備考欄: $remarks";
    $headers = 'From: webmaster@example.com'; // 差出人のメールアドレスを指定

    mail($to, $subject, $message, $headers);

    // エクセルファイルへの保存
    require 'vendor/autoload.php'; // PHPExcelを使用する場合はComposerでPHPExcelをインストールしておく

    $file = 'reports.xlsx'; // 保存するエクセルファイルのパス

    if (file_exists($file)) {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file);
    } else {
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
    }

    $sheet = $spreadsheet->getActiveSheet();
    $row = $sheet->getHighestRow() + 1;

    $sheet->setCellValue('A' . $row, $churchName);
    $sheet->setCellValue('B' . $row, $reporter);
    $sheet->setCellValue('C' . $row, $personName);
    $sheet->setCellValue('D' . $row, $gender);
    $sheet->setCellValue('E' . $row, $ageGroup);
    $sheet->setCellValue('F' . $row, $status);
    $sheet->setCellValue('G' . $row, $address);
    $sheet->setCellValue('H' . $row, $remarks);

    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save($file);

    echo "フォームが正常に送信されました。";
}
?>
