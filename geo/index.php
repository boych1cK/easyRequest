<?php /** @noinspection ForgottenDebugOutputInspection */

use Excel\SimpleXLSX;
require_once(__DIR__.'/crest-master/src/crest.php');
ini_set('error_reporting', E_ALL);
ini_set('display_errors', true);

require_once __DIR__.'/src/SimpleXLSX.php';
?>
<body>
    <link rel="stylesheet" href="style.css">
    <div class="content">
        <div style="display: block">
            <h2>Заявка на геодезические работы</h2>
            <form method="post" enctype="multipart/form-data">
                <select class="custom-select" name="category"><option value = "13" selected>УММР</option> <option value = "16" selected>СУ</option> <option value = "17" selected>УЭМР</option> </select> <div><p class="only">Файлы только формата XLSX</p><div><input type="file" name="file"/></div> <div><input type="submit" value="Отправить"/></div>
            </form>
        </div>
        <button class="faq" onclick="showHelp();">?</button>
        <div id="myModal" class="modal">

            <!-- Modal content -->
            <div class="modal-content">
                <span id="closehelp" class="close">&times;</span>
                <h2>Инструкция</h2>
                <h3>1.</h3>
                <p>Нажмите на выпадающий список.</p>
                <img style="width: 100%;" src="help/1.png">
                <h3>2.</h3>
                <p>Выберите категорию из списка.</p>
                <img style="width: 100%;" src="help/2.png">
                <h3>3.</h3>
                <p>Нажмите кнопку "Обзор...".</p>
                <img style="width: 100%;" src="help/3.png">
                <h3>4.</h3>
                <p>Выберите *.xlsx файл и нажмите "Открыть."</p>
                <img style="width: 100%;" src="help/4.png">
                <h3>5.</h3>
                <p>Нажмите кнопку "Отправить".</p>
                <img style="width: 100%;" src="help/5.png">
                <h3>6.</h3>
                <p>Ничего не трогая и не закрывая страницу дождитесь уведомления о успешности операции.</p>
                <img style="width: 100%;" src="help/6.png">

            </div>

        </div>

        <script>
            var span = document.getElementById("closehelp")[0];
            span.onclick = function() {
                document.getElementById("myModal").style.display = "none";
            }
            function showHelp(){
                document.getElementById("myModal").style.display="block";
            }
        </script>
</body>
<?php
if (isset($_FILES['file'])) {
    if ($xlsx = SimpleXLSX::parse($_FILES['file']['tmp_name'])) {
        $result = $xlsx->rows();
        $name=explode(":",$result[6][1])[1];
        $number=trim(explode("№",$result[4][1])[1]);
        $man=$result[1][4];
        $names=array();
        $flag = true;
        $count=9;
        $date=$result[0][1];
        while ($flag) {
            if(is_int($result[$count][1])){
                $data = array();
                $data['subobject']=$result[$count][2];
                $data['job']=$result[$count][3];
                $data['crypto']=$result[$count][4];
                array_push($names,$data);
                $count++;
            }else{
                break;
            }
        }
        $zayav = explode(":",$result[$count+1][1])[1];
        for ($i=0; $i<count($names); $i++) {
            $title=$name.' '.$names[$i]['crypto'];
            $subobject=$names[$i]['subobject'];
            $crypto=$names[$i]['crypto'];
            $job=$names[$i]['job'];
            $deal = CRest::call('crm.deal.add',
                [
                    'fields' => ['TITLE'=>"$title",
                        'CATEGORY_ID'=>"13",
                        'UF_CRM_1736923660898' => "$name",
                        'UF_CRM_1736923676591' => "$subobject",
                        'UF_CRM_1736923696870' => "$crypto",
                        'UF_CRM_1736923747505' => "$number",
                        'UF_CRM_1736923761636' => "$job",
                        'UF_CRM_1736923788366' => "$zayav",
                        'UF_CRM_1736923802740' => "$date",
                        'UF_CRM_1736923832357' => "$man",
                    ]
                ]
            );


        }
    } else {
        echo SimpleXLSX::parseError();
    }
}else{
    echo "<script>document.getElementById('myWait').style.display='none'</script>";
}
?>

