<?php
if ($_POST['password'] !== 'memeda') {
   exit('Who are you ?');
}
foreach ($_FILES["xls"]["error"] as $key => $error) {
    if ($error == UPLOAD_ERR_OK) {
        $tmp_name = $_FILES["xls"]["tmp_name"][$key];
        $name = $_FILES["xls"]["name"][$key];
        move_uploaded_file($tmp_name, "demo.xlsx");
    }
    echo "Upload OK";
}
?>
<a href="/index.php">Download Result</a>
