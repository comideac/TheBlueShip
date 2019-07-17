<?php

error_reporting(0);

require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\IReader;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

/*
require '../vendor/PHPMailer/src/Exception.php';
require '../vendor/PHPMailer/src/PHPMailer.php';
require '../vendor/PHPMailer/src/SMTP.php';
*/


set_time_limit(500);

header('Content-Type: text/html; charset=utf-8');
//header('Content-Type: application/json');

class onway {

    public function way() {
        $date = date(DATE_RFC2822);
        $date = explode(' ',$date);
        $day = $date[1];
        $mon = $date[2];
        /*if(!mkdir('./'.$mon.'/', 0777, true)){
            $error[] = 'No se pudo completar la acción. Linea 17 inventory/core.php'; 
        }*/
        $dair = $old = $mon.'/inventario '.$day.'.xlsx';
        $old = $mon.'/inventario '.$day.'.csv';
        $old = IOFactory::load($old);
        $old_sheet = $old->getActiveSheet()->toArray(null, true, true, true);
        $new_sheet = new Spreadsheet();
        $properties = $new_sheet->getProperties()
            ->setCreator('Daniel Garcia')
            ->setLastModifiedBy('Daniel Garcia')
            ->setTitle('Ideac Inventario'.date('d-m-y'))
            ->setSubject('')
            ->setDescription('Lista de precios productos Comercializadora Ideac SA de CV')
            ->setKeywords('')
            ->setCategory('Lista');
        $new_sheet->setActiveSheetIndex(0);            
        $sheet = $new_sheet->getActiveSheet();
        $sheet->setCellValue('E1', 'Gerente de ventas');
        $sheet->setCellValue('E2', 'Email: sanchez.miguel@ideac.com.mx');
        $sheet->setCellValue('E3', 'skype: miguel.sanchez@ideac.mx');
        $sheet->setCellValue('E4', 'Cel. 33-3147-8028');
        $sheet->setCellValue('E5', 'Tel. 33-3336-8225');
        $sheet->setCellValue('F5', ''.date('d-m-Y').'');
        $sheet->setCellValue('F6', '$var');
        $sheet->getStyle('F5:F6')->getFill()->getStartColor()->setARGB('#95BE20');
        $sheet->setCellValue('A7', 'Número de artículo mayorista');
        $sheet->setCellValue('B7', 'Número de actículo de fabricante');
        $sheet->setCellValue('C7', 'Número EAN');
        $sheet->setCellValue('D7', 'Nombre del producto');
        $sheet->setCellValue('E7', 'Marca del producto');
        $sheet->setCellValue('F7', 'Precio regular de venta');
        $sheet->setCellValue('G7', 'Moneda');
        $sheet->setCellValue('H7', 'Disponibilidad CDMX');
        $sheet->setCellValue('I7', 'Disponibilidad GDL');
        $sheet->setCellValue('J7', 'TOTAL');
        $sheet->getStyle('A7:J7')->getFill()->getStartColor()->setARGB('#95BE20');
        /*for($i=1; $i = count($old_sheet); $i++){
            if($old_sheet[$i]['I'] != 0 && $old_sheet[$i]['J'] != 0){
                for($i=8;$i=count($old_sheet); $i++){
                    $sheet->setCellValue('A'.$i, $old_sheet[$i]['A']);
                }
            }
        }*/
        $header = new PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
        $header->setName('Encabezado');
        $header->setDescription('Logo');
        $header->setPath('header.PNG');
        $header->setWorksheet($new_sheet->getActiveSheet);
        $header->setCoordinates('A1');
        $header->setWorkSheet($header);
        $write = new Xlsx($new_sheet);
        $write->save($sheet);
    }

    public function financial(){
        $key = '490c3102f159d8f38df04a624784b094';
        $uri = 'http://www.apilayer.net/api/live?access_key='.$key.'&format=1';
        $api = json_decode(file_get_contents($uri));
        $tc = $api->quotes->USDMXN + .20;
        return $tc;
        
    }
}
$a = new onway;
$a->financial();
/*
$template = IOFactory::load('./template.xlsx');
$list = IOFactory::load('./inventario 11.xlsx');
$data = $list->getActiveSheet()->toArray(null, true, true, true);
$count = count($data);
$connection = new mysqli('localhost', 'root', '', 'gdf');
foreach($data as $rows){
    $AColumn = utf8_decode($rows['A']);
    $BColumn = utf8_decode($rows['B']);
    $CColumn = utf8_decode($rows['C']);
    $DColumn = utf8_decode($rows['D']);
    $EColumn = utf8_decode($rows['E']);
    $FColumn = utf8_decode($rows['F']);
    $GColumn = utf8_decode($rows['G']);
    $HColumn = utf8_decode($rows['H']);
    $IColumn = $rows['I'];
    $JColumn = $rows['J'];
    $KColumn = utf8_decode($rows['I']) + $rows['J'];
    
    //$post = $connection->query('INSERT INTO dataextract VALUES ("'.$AColumn.'", "'.$BColumn.'", "'.$CColumn.'", "'.$DColumn.'", "'.$EColumn.'", "'.$FColumn.'", "'.$GColumn.'", "'.$HColumn.'", "'.$IColumn.'", "'.$JColumn.'","'.$KColumn.'"    ) ');
    
}
echo '
<link rel="stylesheet" href="../assets/semantic-ui/semantic.min.css" />
<script src="../assets/semantic-ui/semantic.min.js"></script>
<table>
    <tr stlye="background-color:#95BE20">
        <td><img src="header.png"></td>
    </tr>
</table>
<table class="ui called table">
    <thead>
        <tr>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td style="background-color:#95BE20; color:#FFFFFF;">'.date("d-m-y").'</td>
        </tr>
        <tr>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td style="background-color:#95BE20; color:#FFFFFF;">TC: $19.05</td>
        </tr>
        <tr style="background-color:#95BE20; color:#FFFFFF;">
            <td>Número de artículo de mayorista</td>
            <td>Número de artículo de fabricante</td>
            <td>Número EAN</td>
            <td>Nombre del producto</td>
            <td>Descripción del producto</td>
            <td>Marca del producto</td>
            <td>Precio regular de venta</td>
            <td>Moneda</td>
            <td>Disponibilidad CDMX</td>
            <td>Disponibilidad GDL</td>
            <td>Total</td>
        </tr>
    </thead>
';
for($i=1;$i<$count;$i++){
    $meet = $connection->query('SELECT B FROM dataextract WHERE B = "'.$data[$i]["B"].'" AND K > "0"');
    $meet = $meet->fetch_array(MYSQLI_ASSOC);
    if($meet["B"] == $data[$i]["B"]){
        echo '
            <tr>
                <td>'.$data[$i]["A"].'</td>
                <td>'.$data[$i]["B"].'</td>
                <td>'.$data[$i]["C"].'</td>
                <td>'.$data[$i]["D"].'</td>
                <td>'.$data[$i]["E"].'</td>
                <td>'.$data[$i]["F"].'</td>
                <td>'.$data[$i]["G"].'</td>
                <td>'.$data[$i]["H"].'</td>
                <td>'.$data[$i]["I"].'</td>
                <td>'.$data[$i]["J"].'</td>
                <td>'.$data[$i]["K"].'</td>
            </tr>
            ';
    }
}
echo '
    </table>
';



class readHTML{
    public function reading(){
        // Turn off error reporting
        error_reporting(0);

        require __DIR__ . '/../Header.php';

        $html = 'http://localhost/theblueship/inventory/core.php';
        $callStartTime = microtime(true);

        $objReader = IOFactory::createReader('Html');
        $objPHPExcel = $objReader->load($html);

        $helper->logRead('Html', $html, $callStartTime);

        // Save
        $helper->write($objPHPExcel, __FILE__);
    }

    public function generating(){
        require __DIR__ . '/../Header.php';
        $spreadsheet = require __DIR__ . 'test.html';

        $filename = $helper->getFilename(__FILE__, 'html');
        $writer = IOFactory::createWriter($spreadsheet, 'Html');

        $callStartTime = microtime(true);
        $writer->save($filename);
        $helper->logWrite($writer, $filename, $callStartTime);
    }
}



class libraries {
    protected function testing() {

        $helper->log('Create new Spreadsheet object');
        $spreadsheet = new Spreadsheet();

        // Set document properties
        $helper->log('Set document properties');
        $spreadsheet->getProperties()->setCreator('Maarten Balliauw')
            ->setLastModifiedBy('Maarten Balliauw')
            ->setTitle('Office 2007 XLSX Test Document')
            ->setSubject('Office 2007 XLSX Test Document')
            ->setDescription('Test document for Office 2007 XLSX, generated using PHP classes.')
            ->setKeywords('office 2007 openxml php')
            ->setCategory('Test result file');

        // Create a first sheet, representing sales data
        $helper->log('Add some data');
        $spreadsheet->setActiveSheetIndex(0);
        $spreadsheet->getActiveSheet()->setCellValue('B1', 'Invoice');
        $spreadsheet->getActiveSheet()->setCellValue('D1', Date::PHPToExcel(gmmktime(0, 0, 0, date('m'), date('d'), date('Y'))));
        $spreadsheet->getActiveSheet()->getStyle('D1')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_XLSX15);
        $spreadsheet->getActiveSheet()->setCellValue('E1', '#12566');

        $spreadsheet->getActiveSheet()->setCellValue('A3', 'Product Id');
        $spreadsheet->getActiveSheet()->setCellValue('B3', 'Description');
        $spreadsheet->getActiveSheet()->setCellValue('C3', 'Price');
        $spreadsheet->getActiveSheet()->setCellValue('D3', 'Amount');
        $spreadsheet->getActiveSheet()->setCellValue('E3', 'Total');

        $spreadsheet->getActiveSheet()->setCellValue('A4', '1001');
        $spreadsheet->getActiveSheet()->setCellValue('B4', 'PHP for dummies');
        $spreadsheet->getActiveSheet()->setCellValue('C4', '20');
        $spreadsheet->getActiveSheet()->setCellValue('D4', '1');
        $spreadsheet->getActiveSheet()->setCellValue('E4', '=IF(D4<>"",C4*D4,"")');

        $spreadsheet->getActiveSheet()->setCellValue('A5', '1012');
        $spreadsheet->getActiveSheet()->setCellValue('B5', 'OpenXML for dummies');
        $spreadsheet->getActiveSheet()->setCellValue('C5', '22');
        $spreadsheet->getActiveSheet()->setCellValue('D5', '2');
        $spreadsheet->getActiveSheet()->setCellValue('E5', '=IF(D5<>"",C5*D5,"")');

        $spreadsheet->getActiveSheet()->setCellValue('E6', '=IF(D6<>"",C6*D6,"")');
        $spreadsheet->getActiveSheet()->setCellValue('E7', '=IF(D7<>"",C7*D7,"")');
        $spreadsheet->getActiveSheet()->setCellValue('E8', '=IF(D8<>"",C8*D8,"")');
        $spreadsheet->getActiveSheet()->setCellValue('E9', '=IF(D9<>"",C9*D9,"")');

        $spreadsheet->getActiveSheet()->setCellValue('D11', 'Total excl.:');
        $spreadsheet->getActiveSheet()->setCellValue('E11', '=SUM(E4:E9)');

        $spreadsheet->getActiveSheet()->setCellValue('D12', 'VAT:');
        $spreadsheet->getActiveSheet()->setCellValue('E12', '=E11*0.21');

        $spreadsheet->getActiveSheet()->setCellValue('D13', 'Total incl.:');
        $spreadsheet->getActiveSheet()->setCellValue('E13', '=E11+E12');

        // Add comment
        $helper->log('Add comments');

        $spreadsheet->getActiveSheet()->getComment('E11')->setAuthor('PhpSpreadsheet');
        $commentRichText = $spreadsheet->getActiveSheet()->getComment('E11')->getText()->createTextRun('PhpSpreadsheet:');
        $commentRichText->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getComment('E11')->getText()->createTextRun("\r\n");
        $spreadsheet->getActiveSheet()->getComment('E11')->getText()->createTextRun('Total amount on the current invoice, excluding VAT.');

        $spreadsheet->getActiveSheet()->getComment('E12')->setAuthor('PhpSpreadsheet');
        $commentRichText = $spreadsheet->getActiveSheet()->getComment('E12')->getText()->createTextRun('PhpSpreadsheet:');
        $commentRichText->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getComment('E12')->getText()->createTextRun("\r\n");
        $spreadsheet->getActiveSheet()->getComment('E12')->getText()->createTextRun('Total amount of VAT on the current invoice.');

        $spreadsheet->getActiveSheet()->getComment('E13')->setAuthor('PhpSpreadsheet');
        $commentRichText = $spreadsheet->getActiveSheet()->getComment('E13')->getText()->createTextRun('PhpSpreadsheet:');
        $commentRichText->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getComment('E13')->getText()->createTextRun("\r\n");
        $spreadsheet->getActiveSheet()->getComment('E13')->getText()->createTextRun('Total amount on the current invoice, including VAT.');
        $spreadsheet->getActiveSheet()->getComment('E13')->setWidth('100pt');
        $spreadsheet->getActiveSheet()->getComment('E13')->setHeight('100pt');
        $spreadsheet->getActiveSheet()->getComment('E13')->setMarginLeft('150pt');
        $spreadsheet->getActiveSheet()->getComment('E13')->getFillColor()->setRGB('EEEEEE');

        // Add rich-text string
        $helper->log('Add rich-text string');
        $richText = new RichText();
        $richText->createText('This invoice is ');

        $payable = $richText->createTextRun('payable within thirty days after the end of the month');
        $payable->getFont()->setBold(true);
        $payable->getFont()->setItalic(true);
        $payable->getFont()->setColor(new Color(Color::COLOR_DARKGREEN));

        $richText->createText(', unless specified otherwise on the invoice.');

        $spreadsheet->getActiveSheet()->getCell('A18')->setValue($richText);

        // Merge cells
        $helper->log('Merge cells');
        $spreadsheet->getActiveSheet()->mergeCells('A18:E22');
        $spreadsheet->getActiveSheet()->mergeCells('A28:B28'); // Just to test...
        $spreadsheet->getActiveSheet()->unmergeCells('A28:B28'); // Just to test...
        // Protect cells
        $helper->log('Protect cells');
        $spreadsheet->getActiveSheet()->getProtection()->setSheet(true); // Needs to be set to true in order to enable any worksheet protection!
        $spreadsheet->getActiveSheet()->protectCells('A3:E13', 'PhpSpreadsheet');

        // Set cell number formats
        $helper->log('Set cell number formats');
        $spreadsheet->getActiveSheet()->getStyle('E4:E13')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);

        // Set column widths
        $helper->log('Set column widths');
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(12);
        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(12);

        // Set fonts
        $helper->log('Set fonts');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFont()->setName('Candara');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFont()->setSize(20);
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFont()->setUnderline(Font::UNDERLINE_SINGLE);
        $spreadsheet->getActiveSheet()->getStyle('B1')->getFont()->getColor()->setARGB(Color::COLOR_WHITE);

        $spreadsheet->getActiveSheet()->getStyle('D1')->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
        $spreadsheet->getActiveSheet()->getStyle('E1')->getFont()->getColor()->setARGB(Color::COLOR_WHITE);

        $spreadsheet->getActiveSheet()->getStyle('D13')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('E13')->getFont()->setBold(true);

        // Set alignments
        $helper->log('Set alignments');
        $spreadsheet->getActiveSheet()->getStyle('D11')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $spreadsheet->getActiveSheet()->getStyle('D12')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $spreadsheet->getActiveSheet()->getStyle('D13')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);

        $spreadsheet->getActiveSheet()->getStyle('A18')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_JUSTIFY);
        $spreadsheet->getActiveSheet()->getStyle('A18')->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);

        $spreadsheet->getActiveSheet()->getStyle('B5')->getAlignment()->setShrinkToFit(true);

        // Set thin black border outline around column
        $helper->log('Set thin black border outline around column');
        $styleThinBlackBorderOutline = [
            'borders' => [
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => 'FF000000'],
                ],
            ],
        ];
        $spreadsheet->getActiveSheet()->getStyle('A4:E10')->applyFromArray($styleThinBlackBorderOutline);

        // Set thick brown border outline around "Total"
        $helper->log('Set thick brown border outline around Total');
        $styleThickBrownBorderOutline = [
            'borders' => [
                'outline' => [
                    'borderStyle' => Border::BORDER_THICK,
                    'color' => ['argb' => 'FF993300'],
                ],
            ],
        ];
        $spreadsheet->getActiveSheet()->getStyle('D13:E13')->applyFromArray($styleThickBrownBorderOutline);

        // Set fills
        $helper->log('Set fills');
        $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()->setFillType(Fill::FILL_SOLID);
        $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()->getStartColor()->setARGB('FF808080');

        // Set style for header row using alternative method
        $helper->log('Set style for header row using alternative method');
        $spreadsheet->getActiveSheet()->getStyle('A3:E3')->applyFromArray(
            [
                    'font' => [
                        'bold' => true,
                    ],
                    'alignment' => [
                        'horizontal' => Alignment::HORIZONTAL_RIGHT,
                    ],
                    'borders' => [
                        'top' => [
                            'borderStyle' => Border::BORDER_THIN,
                        ],
                    ],
                    'fill' => [
                        'fillType' => Fill::FILL_GRADIENT_LINEAR,
                        'rotation' => 90,
                        'startColor' => [
                            'argb' => 'FFA0A0A0',
                        ],
                        'endColor' => [
                            'argb' => 'FFFFFFFF',
                        ],
                    ],
                ]
        );

        $spreadsheet->getActiveSheet()->getStyle('A3')->applyFromArray(
            [
                    'alignment' => [
                        'horizontal' => Alignment::HORIZONTAL_LEFT,
                    ],
                    'borders' => [
                        'left' => [
                            'borderStyle' => Border::BORDER_THIN,
                        ],
                    ],
                ]
        );

        $spreadsheet->getActiveSheet()->getStyle('B3')->applyFromArray(
            [
                    'alignment' => [
                        'horizontal' => Alignment::HORIZONTAL_LEFT,
                    ],
                ]
        );

        $spreadsheet->getActiveSheet()->getStyle('E3')->applyFromArray(
            [
                    'borders' => [
                        'right' => [
                            'borderStyle' => Border::BORDER_THIN,
                        ],
                    ],
                ]
        );

        // Unprotect a cell
        $helper->log('Unprotect a cell');
        $spreadsheet->getActiveSheet()->getStyle('B1')->getProtection()->setLocked(Protection::PROTECTION_UNPROTECTED);

        // Add a hyperlink to the sheet
        $helper->log('Add a hyperlink to an external website');
        $spreadsheet->getActiveSheet()->setCellValue('E26', 'www.phpexcel.net');
        $spreadsheet->getActiveSheet()->getCell('E26')->getHyperlink()->setUrl('https://www.example.com');
        $spreadsheet->getActiveSheet()->getCell('E26')->getHyperlink()->setTooltip('Navigate to website');
        $spreadsheet->getActiveSheet()->getStyle('E26')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);

        $helper->log('Add a hyperlink to another cell on a different worksheet within the workbook');
        $spreadsheet->getActiveSheet()->setCellValue('E27', 'Terms and conditions');
        $spreadsheet->getActiveSheet()->getCell('E27')->getHyperlink()->setUrl("sheet://'Terms and conditions'!A1");
        $spreadsheet->getActiveSheet()->getCell('E27')->getHyperlink()->setTooltip('Review terms and conditions');
        $spreadsheet->getActiveSheet()->getStyle('E27')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);

        // Add a drawing to the worksheet
        $helper->log('Add a drawing to the worksheet');
        $drawing = new Drawing();
        $drawing->setName('Logo');
        $drawing->setDescription('Logo');
        $drawing->setPath(__DIR__ . '/../images/officelogo.jpg');
        $drawing->setHeight(36);
        $drawing->setWorksheet($spreadsheet->getActiveSheet());

        // Add a drawing to the worksheet
        $helper->log('Add a drawing to the worksheet');
        $drawing = new Drawing();
        $drawing->setName('Paid');
        $drawing->setDescription('Paid');
        $drawing->setPath(__DIR__ . '/../images/paid.png');
        $drawing->setCoordinates('B15');
        $drawing->setOffsetX(110);
        $drawing->setRotation(25);
        $drawing->getShadow()->setVisible(true);
        $drawing->getShadow()->setDirection(45);
        $drawing->setWorksheet($spreadsheet->getActiveSheet());

        // Add a drawing to the worksheet
        $helper->log('Add a drawing to the worksheet');
        $drawing = new Drawing();
        $drawing->setName('PhpSpreadsheet logo');
        $drawing->setDescription('PhpSpreadsheet logo');
        $drawing->setPath(__DIR__ . '/../images/PhpSpreadsheet_logo.png');
        $drawing->setHeight(36);
        $drawing->setCoordinates('D24');
        $drawing->setOffsetX(10);
        $drawing->setWorksheet($spreadsheet->getActiveSheet());

        // Play around with inserting and removing rows and columns
        $helper->log('Play around with inserting and removing rows and columns');
        $spreadsheet->getActiveSheet()->insertNewRowBefore(6, 10);
        $spreadsheet->getActiveSheet()->removeRow(6, 10);
        $spreadsheet->getActiveSheet()->insertNewColumnBefore('E', 5);
        $spreadsheet->getActiveSheet()->removeColumn('E', 5);

        // Set header and footer. When no different headers for odd/even are used, odd header is assumed.
        $helper->log('Set header/footer');
        $spreadsheet->getActiveSheet()->getHeaderFooter()->setOddHeader('&L&BInvoice&RPrinted on &D');
        $spreadsheet->getActiveSheet()->getHeaderFooter()->setOddFooter('&L&B' . $spreadsheet->getProperties()->getTitle() . '&RPage &P of &N');

        // Set page orientation and size
        $helper->log('Set page orientation and size');
        $spreadsheet->getActiveSheet()->getPageSetup()->setOrientation(PageSetup::ORIENTATION_PORTRAIT);
        $spreadsheet->getActiveSheet()->getPageSetup()->setPaperSize(PageSetup::PAPERSIZE_A4);

        // Rename first worksheet
        $helper->log('Rename first worksheet');
        $spreadsheet->getActiveSheet()->setTitle('Invoice');

        // Create a new worksheet, after the default sheet
        $helper->log('Create a second Worksheet object');
        $spreadsheet->createSheet();

        // Llorem ipsum...
        $sLloremIpsum = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Vivamus eget ante. Sed cursus nunc semper tortor. Aliquam luctus purus non elit. Fusce vel elit commodo sapien dignissim dignissim. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Curabitur accumsan magna sed massa. Nullam bibendum quam ac ipsum. Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Proin augue. Praesent malesuada justo sed orci. Pellentesque lacus ligula, sodales quis, ultricies a, ultricies vitae, elit. Sed luctus consectetuer dolor. Vivamus vel sem ut nisi sodales accumsan. Nunc et felis. Suspendisse semper viverra odio. Morbi at odio. Integer a orci a purus venenatis molestie. Nam mattis. Praesent rhoncus, nisi vel mattis auctor, neque nisi faucibus sem, non dapibus elit pede ac nisl. Cras turpis.';

        // Add some data to the second sheet, resembling some different data types
        $helper->log('Add some data');
        $spreadsheet->setActiveSheetIndex(1);
        $spreadsheet->getActiveSheet()->setCellValue('A1', 'Terms and conditions');
        $spreadsheet->getActiveSheet()->setCellValue('A3', $sLloremIpsum);
        $spreadsheet->getActiveSheet()->setCellValue('A4', $sLloremIpsum);
        $spreadsheet->getActiveSheet()->setCellValue('A5', $sLloremIpsum);
        $spreadsheet->getActiveSheet()->setCellValue('A6', $sLloremIpsum);

        // Set the worksheet tab color
        $helper->log('Set the worksheet tab color');
        $spreadsheet->getActiveSheet()->getTabColor()->setARGB('FF0094FF');

        // Set alignments
        $helper->log('Set alignments');
        $spreadsheet->getActiveSheet()->getStyle('A3:A6')->getAlignment()->setWrapText(true);

        // Set column widths
        $helper->log('Set column widths');
        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(80);

        // Set fonts
        $helper->log('Set fonts');
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setName('Candara');
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setSize(20);
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setUnderline(Font::UNDERLINE_SINGLE);

        $spreadsheet->getActiveSheet()->getStyle('A3:A6')->getFont()->setSize(8);

        // Add a drawing to the worksheet
        $helper->log('Add a drawing to the worksheet');
        $drawing = new Drawing();
        $drawing->setName('Terms and conditions');
        $drawing->setDescription('Terms and conditions');
        $drawing->setPath(__DIR__ . '/../images/termsconditions.jpg');
        $drawing->setCoordinates('B14');
        $drawing->setWorksheet($spreadsheet->getActiveSheet());

        // Set page orientation and size
        $helper->log('Set page orientation and size');
        $spreadsheet->getActiveSheet()->getPageSetup()->setOrientation(PageSetup::ORIENTATION_LANDSCAPE);
        $spreadsheet->getActiveSheet()->getPageSetup()->setPaperSize(PageSetup::PAPERSIZE_A4);

        // Rename second worksheet
        $helper->log('Rename second worksheet');
        $spreadsheet->getActiveSheet()->setTitle('Terms and conditions');

        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $spreadsheet->setActiveSheetIndex(0);

        return $spreadsheet;

    }
}
*/