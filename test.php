<?Php
    #header('Content-Type: application/json');
    //serverName\instanceName, portNumber (por defecto es 1433)
    $serverName = "SENDBOXSERVER\\COMPAC2"; 
    $connectionInfo = array( "Database"=>"adCOMERCIALIZADORAIDE");
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
            echo '<table width="970" border="1">';
            echo '<tr><td width="600">'.utf8_encode($getReason['CRAZONSOCIAL']) . '</td><td width="370">' . utf8_encode($getMail['CEMAIL']) . "</td></tr>";
            echo '</table>';
            #var_dump(utf8_encode($test['CRAZONSOCIAL']));
        }
   }else{
        echo "Connection could not be established.<br />";
        die( print_r( sqlsrv_errors(), true));
   }




