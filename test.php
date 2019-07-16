<?Php

// set up basic connection 
$ftp_server = "ftp.comideac.com"; 
$conn_id = ftp_ssl_connect($ftp_server); 

// login with username and password 
$ftp_user_name = "garcia.daniel@ideac.com.mx"; 
$ftp_user_pass = "zWab!Llz-_j)"; 
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

if (ftp_chdir($conn_id, "ideac.com.mx/inventario")) {
    echo "\nCurrent directory is now: " . ftp_pwd($conn_id) . "\n";
    if (ftp_get($conn_id, 'ideac '.date('d-m-y').'.xlsx', 'inventario.csv', FTP_BINARY)) {
        echo "Se ha guardado satisfactoriamente en ideac ".date('d-m-y').".xlsx\n";
    } else {
        echo "Ha habido un problema\n";
    }
} else { 
    echo "Couldn't change directory\n";
}

$buff = ftp_rawlist($conn_id, '.'); 
//var_dump($buff); 
ftp_close($conn_id);  




//recipient
$to = 'daniel4581@protonmail.com';

//sender
$from = 'garcia.daniel@ideac.com.mx';
$fromName = 'Daniel Garcia | Comercializadora Ideac SA. de CV.';

//email subject
$subject = 'Inventario '.date('d-m-y'); 

//attachment file path
$file = "ideac ".date('d-m-y').'xlsx';

//email body content
$htmlContent = '<h1>Inventario</h1>'.date('d-m-y').'<p>Cualquier duda estoy a tus ordenes. Que tengas excelente día</p><br/><br/><p>saludos</p><br /><p>Daniel Garcia</p>';

//header for sender info
$headers = "From: $fromName"." <".$from.">";

//boundary 
$semi_rand = md5(time()); 
$mime_boundary = "==Multipart_Boundary_x".$semi_rand."x"; 

//headers for attachment 
$headers .= "nMIME-Version: 1.0n Content-Type: multipart/mixed;n boundary='".$mime_boundary."'"; 

//multipart boundary 
$message = "--".$mime_boundary."n" . "Content-Type: text/html; charset='UTF-8'n" .
"Content-Transfer-Encoding: 7bitnn" . $htmlContent . "nn"; 

//preparing attachment
if(!empty($file) > 0){
    if(is_file($file)){
        $message .= "--".$mime_boundary."n";
        $fp =    @fopen($file,"rb");
        $data =  @fread($fp,filesize($file));

        @fclose($fp);
        $data = chunk_split(base64_encode($data));
        $message .= "Content-Type: application/octet-stream; name='".basename($file)."'n" . 
        "Content-Description: '".basename($files[$i])."'n" .
        "Content-Disposition: attachment;n filename='".basename($file)."'; size='".filesize($file)."';n" . 
        "Content-Transfer-Encoding: base64nn" . $data . "nn";
    }
}
$message .= "--".$mime_boundary."--";
$returnpath = "-f" . $from;

//send email
$mail = @mail($to, $subject, $message, $headers, $returnpath); 
var_dump($mail);

//email sending status
echo $mail?"<h1>Mail sent.</h1>":"<h1>Mail sending.</h1>";



/* ----------------------------------------------------------- */

///////Configuración/////
$mail_destinatario = 'daniel4581@protonmail.com';
///////Fin configuración//

///// Funciones necesarias////
function form_mail($sPara, $sAsunto, $sTexto, $sDe){
    $bHayFicheros = 0;
    $sCabeceraTexto = "";
    $sAdjuntos = "";
    if ($sDe)$sCabeceras = "From:".$sDe."\n";
    else $sCabeceras = "";
    $sCabeceras .= "MIME-version: 1.0\n";
    foreach ($_POST as $sNombre => $sValor)
        $sTexto = $sTexto."\n".$sNombre." = ".$sValor;
        foreach ($_FILES as $vAdjunto){
            if ($bHayFicheros == 0){
                $bHayFicheros = 1;
                $sCabeceras .= "Content-type: multipart/mixed;";
                $sCabeceras .= "boundary=\"--_Separador-de-mensajes_--\"\n";
                $sCabeceraTexto = "----_Separador-de-mensajes_--\n";
                $sCabeceraTexto .= "Content-type: text/plain;charset=iso-8859-1\n";
                $sCabeceraTexto .= "Content-transfer-encoding: 7BIT\n";
                $sTexto = $sCabeceraTexto.$sTexto;
            }
            if ($vAdjunto["size"] > 0) {
                $sAdjuntos .= "\n\n----_Separador-de-mensajes_--\n";
                $sAdjuntos .= "Content-type: ".$vAdjunto["type"].";name=\"".$vAdjunto["name"]."\"\n";;
                $sAdjuntos .= "Content-Transfer-Encoding: BASE64\n";
                $sAdjuntos .= "Content-disposition: attachment;filename=\"".$vAdjunto["name"]."\"\n\n";
                $oFichero = fopen($vAdjunto["tmp_name"], 'r');
                $sContenido = fread($oFichero, filesize($vAdjunto["tmp_name"]));
                $sAdjuntos .= chunk_split(base64_encode($sContenido));
                fclose($oFichero);
            }
        }
        if ($bHayFicheros)
            $sTexto .= $sAdjuntos."\n\n----_Separador-de-mensajes_----\n";
            return(mail($sPara, $sAsunto, $sTexto, $sCabeceras));
        }

        if (isset ($_POST['enviar'])) {
            if (form_mail($mail_destinatario, $_POST['asunto'],
                "Los datos introducidos en el formulario son:\n\n", $_POST['email']))
                echo 'Su mensaje a sido enviado correctamente. Gracias por contactar con nosostros';
            else echo 'Error al enviar el formulario. Por favor, inténtelo de nuevo mas tarde.'; 
        }

form_mail('daniel4581@protonmail.com', 'Inventario', 'Te envio el inventario', 'garcia.daniel@ideac.com.mx');