<?php

require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\IReader;
use PhpOffice\PhpSpreadsheet\Style\Font;

$file = 'inventario 12.xlsx';
$sheet = IOFactory::load($file);
$sheet = $sheet->getActiveSheet()->toArray(null, true, true, true);

echo 
'
    <table>
        <tr>
            <td style="background-color:#95BE20">'.$sheet[1]["A"].'</td>
            <td style="background-color:#95BE20">'.$sheet[1]["B"].'</td>
            <td style="background-color:#95BE20">'.$sheet[1]["C"].'</td>
            <td style="background-color:#95BE20">'.$sheet[1]["D"].'</td>
            <td style="background-color:#95BE20">'.$sheet[1]["E"].'</td>
            <td style="background-color:#95BE20">'.$sheet[1]["F"].'</td>
            <td style="background-color:#95BE20">'.$sheet[1]["G"].'</td>
            <td style="background-color:#95BE20">'.$sheet[1]["H"].'</td>
            <td style="background-color:#95BE20">'.$sheet[1]["I"].'</td>
            <td style="background-color:#95BE20">'.$sheet[1]["J"].'</td>
        </tr>
        <tr>
            <td style="background-color:#FFFFFF">'.$sheet[2]["A"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[2]["B"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[2]["C"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[2]["D"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[2]["E"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[2]["F"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[2]["G"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[2]["H"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[2]["I"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[2]["J"].'</td>
        </tr>
        <tr>
            <td style="background-color:#FFFFFF">'.$sheet[3]["A"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[3]["B"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[3]["C"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[3]["D"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[3]["E"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[3]["F"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[3]["G"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[3]["H"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[3]["I"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[3]["J"].'</td>
        </tr>
        <tr>
            <td style="background-color:#FFFFFF">'.$sheet[4]["A"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[4]["B"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[4]["C"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[4]["D"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[4]["E"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[4]["F"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[4]["G"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[4]["H"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[4]["I"].'</td>
            <td style="background-color:#FFFFFF">'.$sheet[4]["J"].'</td>
        </tr>
    
    </table>
';

?>