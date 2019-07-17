<?php

/*class blacklist {
    public function _upload(){
        $dir = 'inventory/';
        $fil = $dir . basename($_FILES['monkey']['name']);
        $ok = 1;
        $fileType = strtolower(pathinfo($fil, PATHINFO_EXTENSION));
        move_uploaded_file($_FILES['monkey']['tmp_name'], $fil);
    }
}
$a = new blacklist;
$a->_upload();
*/
?>

<!DOCTYPE html>
<html>
    <head>
        <title>Blackship</title>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta http-equiv="X-UA-Compatible" content="ie=edge">
        <script src="https://semantic-ui.com/javascript/library/jquery.min.js"></script>
        <script src="assets/semantic-ui/semantic.min.js"></script>
        <link rel="stylesheet" type="text/css" href="assets/semantic-ui/semantic.css" />
        <link href="https://fonts.googleapis.com/css?family=Barlow+Condensed&display=swap" rel="stylesheet">
        <style>
            * {
                margin:0;
                padding:0;
            }
            body {
                font-family: 'Barlow Condensed', sans-serif;
            }
            .header {
                width:100%;
                text-align: center;
                padding-top:2%;
                padding-bottom:2%;
            }
            .content {
                width:50%;
                display:block;
                margin-left:auto;
                margin-right:auto;
            }
            #shiptr.hiden {
                overflow: hidden;
            }
            .footer {
                text-align: center;
                margin-top:28%;
            }
        </style>
    </head>
    <body>
        <div class="header">
            <h1>The Blue Ship</h1>
        </div>
        <div class="content">
            <div class="ui placeholder segment">
                <div class="ui two column stackable center aligned grid">
                    <div class="ui vertical divider">   Ideac</div>
                        <div class="middle aligned row">
                            <div class="column">
                                <div class="ui icon header">
                                    <i class="money bill alternate outline icon"></i>
                                    Tipo de cambio
                                </div>
                                <div class="field">
                                    <div class="ui right labeled input">
                                        <label for="amount" class="ui label">$</label>
                                        <input type="text" placeholder="Ej: $19.20" id="amount">
                                        <div class="ui basic label">Pesos</div>
                                    </div>
                                </div>
                            </div>
                            <div class="column">
                                <div class="ui icon header">
                                    <i class="file excel outline icon"></i>
                                    Cargar archivo Excel
                                </div>
                                    <form enctype="multipart/form-data" action="" method="post" id="shipper">
                                        <input type="file" name="monkey" class="subir ui primary button disabled" onchange="this.form.submit()">
                                    </form>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <div class="footer">
            <p>Hecho para Comercializadora Ideac SA de CV por Daniel Garcia</p>
        </div>
        <script>
            var tocl = $('#amount');
            $('#amount').keypress(function( event ) {
                if (tocl.val().length > 3) {
                    $('.subir').removeClass('disabled');
                }else{
                    $('.subir').addClass('disabled');
                }
            });
        </script>
        <script>
            $('shipper').ajaxForm({url: '/' type: 'post'});
        </script>

            <?php
                $in = file_gets_contents('inventory/template.html');
                $fo = fopen('inventory/ideac '.date('d-m-y'), 'w+');
                $fw = fwrite($fo, $in);
                $fc = fclose($fo);
            ?>

    </body>
</html>