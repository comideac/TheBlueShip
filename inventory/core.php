<?php

require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\IReader;
use \PhpOffice\PhpSpreadsheet\Style\Font;

class onway {

    function way() {
        $date = date(DATE_RFC2822);
        $date = explode(' ',$date);
        $day = $date[1];
        $mon = $date[2];
        /*if(!mkdir('./'.$mon.'/', 0777, true)){
            $error[] = 'No se pudo completar la acción. Linea 17 inventory/core.php'; 
        }*/
        $old = $mon.'/inventario '.$day.'.xlsx';
        $old = IOFactory::load($old);
        $old_sheet = $old->getActiveSheet()->toArray(null, true, true, true);
        $header = new PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
        $header->setName('Encabezado');
        $header->setDescription('Logo');
        $header->setPath('../assets/header.PNG');
        $header->setCoordinates('A1');
        $new_sheet = new Spreadsheet();
        $properties = $new_sheet->getProperties()
            ->setCreator('Daniel Garcia')
            ->setLastModifiedBy('Daniel Garcia')
            ->setTitle('Ideac Inventario'.date('d-m-y'))
            ->setSubject('')
            ->setDescription('Lista de precios productos Comercializadora Ideac SA de CV')
            ->setKeywords('')
            ->setCategory('Lista');
        $sheet = $new_sheet->getActiveSheet();
        $sheet->setCellValue('E1', 'Gerente de ventas');
        $sheet->setCellValue('E2', 'Email: sanchez.miguel@ideac.com.mx');
        $sheet->setCellValue('E3', 'skype: miguel.sanchez@ideac.mx');
        $sheet->setCellValue('E4', 'Cel. 33-3147-8028');
        $sheet->setCellValue('E5', 'Tel. 33-3336-8225');
        $sheet->setCellValue('F5', date('d-m-Y'));
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
        
        $writer = new Xlsx();
        $writer->save('text.xlsx');
            /*if($old_sheet[$i]['I'] == 0 && $old_sheet[$i]['J'] == 0){
                
                $l = json_encode($old_sheet);
                var_dump(json_decode($l));
            }*/
    }
}
$a = new onway;
$a->way();