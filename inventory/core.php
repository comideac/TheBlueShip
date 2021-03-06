<?php

#ini_set('memory_limit', '-1');

require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\IReader;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column;
use PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

require '../vendor/PHPMailer/PHPMailer/src/Exception.php';
require '../vendor/PHPMailer/PHPMailer/src/PHPMailer.php';
require '../vendor/PHPMailer/PHPMailer/src/SMTP.php';

set_time_limit(500);

header('Content-Type: text/html; charset=utf-8');
#header('Content-Type: application/json');

class onway {

    public function eftipi(){
        // set up basic connection 
        $ftp_server = "ftp.**.com"; 
        $conn_id = ftp_ssl_connect($ftp_server); 

        // login with username and password 
        $ftp_user_name = "**"; 
        $ftp_user_pass = "**"; 
        $login_result = ftp_login($conn_id, $ftp_user_name, $ftp_user_pass); 
        ftp_pasv($conn_id, true); 
        // check connection 
        if ((!$conn_id) || (!$login_result)) { 
            echo "FTP connection has failed!"; 
            echo "Attempted to connect to $ftp_server for user $ftp_user_name"; 
            exit; 
        } else { 
            echo "Connected to $ftp_server, for user $ftp_user_name"; 
        } 

        if (ftp_chdir($conn_id, "**mx/inventario")) {
            echo "\nCurrent directory is now: " . ftp_pwd($conn_id) . "\n";
            if (ftp_get($conn_id, 'ideac '.date('d-m-y').'.csv', 'inventario.csv', FTP_BINARY)) {
                echo "Se ha guardado satisfactoriamente en 'ideac ".date('d-m-y').".csv'\n";
                rename($_SERVER['DOCUMENT_ROOT'].'/TheBlueShip/inventory/ideac '.date('d-m-y').'.csv', $_SERVER['DOCUMENT_ROOT'].'/TheBlueShip/inventory/Jul/ideac '.date('d-m-y').'.csv');
            } else {
                echo "Ha habido un problema\n";
            }
        } else { 
            echo "Couldn't change directory\n";
        }

        $buff = ftp_rawlist($conn_id, '.'); 
        //var_dump($buff); 
        ftp_close($conn_id);  
    }

    public function financial(){
        $key = '**';
        $uri = 'http://www.apilayer.net/api/live?access_key='.$key.'&format=1';
        $api = json_decode(file_get_contents($uri));
        $tc = $api->quotes->USDMXN + .07;
        return $tc;
    }

    public function way() {
        $date = date(DATE_RFC2822);
        $date = explode(' ',$date);
        $day = $date[1];
        $mon = $date[2];

        $dair = $old = $mon.'/inventario '.$day.'.xlsx';
        $old = $mon.'/ideac '.date('d-m-y').'.csv';
        $old = IOFactory::load($old);
        $old_sheet = $old->getActiveSheet()->toArray(null, true, true, true);
        $maxRow = $old->getActiveSheet()->getHighestRow();
        $maxCol = $old->getActiveSheet()->getHighestColumn();
        #$old_sheet = $old_sheet->rangeToArray('A1:' . $maxCol . $maxRow);
        $new_sheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $txt = $new_sheet;
        $properties = $new_sheet->getProperties()
            ->setCreator('Daniel Garcia')
            ->setLastModifiedBy('Daniel Garcia')
            ->setTitle('Ideac Inventario'.date('d-m-y'))
            ->setSubject('')
            ->setDescription('Lista de precios productos Comercializadora Ideac SA de CV')
            ->setKeywords('')
            ->setCategory('Lista');
        $sheet = $new_sheet->getActiveSheet();
        $sheet->getColumnDimension('A')->setWidth(30);
        $sheet->getColumnDimension('B')->setWidth(33);
        $sheet->getColumnDimension('C')->setWidth(17);
        $sheet->getColumnDimension('D')->setWidth(70);
        $sheet->getColumnDimension('E')->setWidth(21);
        $sheet->getColumnDimension('F')->setWidth(25);
        $sheet->getColumnDimension('G')->setWidth(11);
        $sheet->getColumnDimension('H')->setWidth(22);
        $sheet->getColumnDimension('I')->setWidth(21);
        $sheet->getColumnDimension('J')->setWidth(11);
        $sheet->setCellValue('F5', date('d/m/y'));
        $sheet->setCellValue('F6', 'TC: '.self::financial());
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
        $sheet->getStyle('A7:J7')
            ->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE);
        $sheet->getStyle('A7:J7')->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()->setARGB('95BE20');

        $sheet->getStyle('A1:Q6')->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()->setARGB('FFFFFF');

        for($i=8;$i < count($old_sheet); $i++){
            $test = $old_sheet[$i]['A'];
            $test1 = array($i => $test);
        
            $CB = $old_sheet[$i]['B'];
            $CB1 = array($i => $CB); 

            $CC = $old_sheet[$i]['C'];
            $CC1 = array($i => $CC);

            $CD = $old_sheet[$i]['D'];
            $CD1 = array($i => $CD);

            $CF = $old_sheet[$i]['F'];
            $CF1 = array($i => $CF);

            $CG = $old_sheet[$i]['G'];
            $CG1 = array($i => $CG);

            $CH = $old_sheet[$i]['H'];  #whore
            $CH1 = array($i => $CH);     #whore

            $CI = $old_sheet[$i]['I'];
            $CI1 = array($i => $CI);

            $CJ = $old_sheet[$i]['J'];
            $CJ1 = array($i => $CJ);

            $CK = $CI + $CJ;

            $ti = $sheet->fromArray($test1, NULL, 'A'.$i);
            $t1 = $sheet->fromArray($CB1, NULL, 'B'.$i);
            $t2 = $sheet->fromArray($CC1, NULL, 'C'.$i);
            $sheet->getStyle('C8:C399')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);
            $sheet->getStyle('C8:C339')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);

            $t3 = $sheet->fromArray($CD1, NULL, 'D'.$i);
            $t4 = $sheet->fromArray($CF1, NULL, 'E'.$i);
            $t5 = $sheet->fromArray($CG1, NULL, 'F'.$i);
            $t6 = $sheet->fromArray(array('A' => 'Dólar'), NULL, 'G'.$i);

            $t7 = $sheet->fromArray($CI1, NULL, 'H'.$i);
            $t8 = $sheet->fromArray($CJ1, NULL, 'I'.$i);
            $ide = $sheet->setCellValue('AD'.$i, $i);
            $t9 = $sheet->setCellValue('J'.$i, '=H'. $i . ' + I' . $i);
            $sheet->getCell('J'.$i)->getStyle()->setQuotePrefix(true);
        }
        $header = new PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
        $header->setName('Encabezado');
        $header->setDescription('Logo');
        $header->setPath('header.PNG');
        $header->setWorksheet($txt->getActiveSheet());
        $header->setCoordinates('A1');
        $write = new Xlsx($new_sheet);
        $write->save($dair);
    }
}


class shipper {
    public function x_000(){

        $serverName = "**\\**"; 
        $connectionInfo = array( "Database"=>"**");
        $conn = sqlsrv_connect( $serverName, $connectionInfo);
        
        if( $conn ) {
            $test = sqlsrv_query($conn, 'SELECT * FROM dbo.admClientes', array(), array("Scrollable" => SQLSRV_CURSOR_KEYSET));
            $test1 = sqlsrv_num_rows($test);
            var_dump($test1);
            for($i=0;$i < $test1; $i++){
                $getReason = sqlsrv_query($conn, 'SELECT * FROM dbo.admClientes WHERE CIDCLIENTEPROVEEDOR = '.$i.'', array(), array("Scrollable" => SQLSRV_CURSOR_KEYSET));
                $getReason = sqlsrv_fetch_array($getReason, SQLSRV_FETCH_ASSOC);
                $getMail = sqlsrv_query($conn, 'SELECT * FROM dbo.admDomicilios WHERE CIDCATALOGO = '.$i.'');
                $getMail = sqlsrv_fetch_array($getMail, SQLSRV_FETCH_ASSOC);
                $mailingList = utf8_encode($getMail['CEMAIL']);
            }
        }else{
            echo "Connection could not be established.<br />";
            die( print_r( sqlsrv_errors(), true));
        }

        $mail = new PHPMailer(true);
        try {
            //Server settings
            $mail->SMTPDebug = 2;                                       // Enable verbose debug output
            $mail->isSMTP();                                            // Set mailer to use SMTP
            $mail->Host       = '**.**.com;**.**.com';    // Specify main and backup SMTP servers
            $mail->SMTPAuth   = true;                                   // Enable SMTP authentication
            $mail->Username   = '**.**@**.**.mx';           // SMTP username
            $mail->Password   = '**';                         // SMTP password
            $mail->SMTPSecure = 'ssl';                                  // Enable TLS encryption, `ssl` also accepted
            $mail->Port       = 465;                                    // TCP port to connect to
            var_dump($mail);
        
            //Recipients
            $mail->setFrom('**.**@**.**.mx', '** **');
            #$mail->addAddress('sanchez.miguel@ideac.com.mx', 'To Daniel Garcia');     // Add a recipient
            $mail->addAddress('**@**.es');               // Name is optional
            #$mail->addReplyTo('info@example.com', 'Information');
            #$mail->addCC('cc@example.com');
            $mail->addBCC('**@**.com');
        
            // Attachments
            $mail->addAttachment('./Jul/ideac '.date('d-m-y').'.csv');         // Add attachments
            #$mail->addAttachment('/tmp/image.jpg', 'new.jpg');    // Optional name
        
            // Content
            $mail->isHTML(true);                                  // Set email format to HTML
            $mail->Subject = 'Inventario '.date('d-m-y');
            $mail->Body    = '<img src="http://comideac.com/bodymail.png">';
            #$mail->AltBody = 'This is the body in plain text for non-HTML mail clients';
        
            $mail->send();
            echo 'Message has been sent';
        } catch (Exception $e) {
            echo "Message could not be sent. Mailer Error: {$mail->ErrorInfo}";
        }
    }
}
$a = new onway;
$a->eftipi();
$a->way();
echo $a->financial();
$b = new shipper;
$b->x_000();
