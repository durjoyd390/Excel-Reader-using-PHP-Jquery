<?php
require 'vendor/autoload.php';

if (isset($_POST['ExcelDataStepA'])) {
if (isset($_FILES['ExcelFile'])) {
$error = '1';

$filename_icon = $_FILES["ExcelFile"]["name"];
$tempname_icon = $_FILES["ExcelFile"]["tmp_name"];
$file_basename_icon = substr($filename_icon, 0, strripos($filename_icon, '.'));
$file_ext_icon = substr($filename_icon, strripos($filename_icon, '.'));
if ($file_ext_icon != '.xlsx' && $file_ext_icon != '.xls'){
$res = 'Your file type is not supported!';
}
else{

$error = '0';
if ($file_ext_icon == '.xlsx') {
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
}
elseif ($file_ext_icon == '.xls') {
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
}


$spreadsheet = $reader->load($tempname_icon);
$worksheetNames = $spreadsheet->getSheetNames();
$res = '<option value="0">Select Sheet</option>';
foreach ($worksheetNames as $sheet) {
$res .= '<option value="'.$sheet.'">'.$sheet.'</option>';
}

}



$return = json_encode(array('error' => $error, 'res' => $res));
exit($return);
}
}
// ---------------------------- End ExcelDataStepA ---------------------------
if (isset($_POST['ExcelDataStepB'])) {
if (isset($_FILES['ExcelFile'])) {
$error = '1';

$filename_icon = $_FILES["ExcelFile"]["name"];
$tempname_icon = $_FILES["ExcelFile"]["tmp_name"];
$file_basename_icon = substr($filename_icon, 0, strripos($filename_icon, '.'));
$file_ext_icon = substr($filename_icon, strripos($filename_icon, '.'));
if ($file_ext_icon != '.xlsx' && $file_ext_icon != '.xls'){
$res = 'Your file type is not supported!';
}
else if ($_POST['ExcelDataStepB'] == 0) {
$res = 'Please Select WorkSheet!';
}
else{
$error = '0';
if ($file_ext_icon == '.xlsx') {
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
}
elseif ($file_ext_icon == '.xls') {
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
}

$spreadsheet = $reader->load($tempname_icon);
$worksheet = $spreadsheet->getSheetByName($_POST['ExcelDataStepB']);
$columnNames = $worksheet->getRowIterator()->current()->getCellIterator();
$res = '<option value="0">Select Column</option>';
foreach ($columnNames as $column) {
$res .= '<option value="'.$column->getColumn().'">'.$column->getValue().'</option>';
}

}



$return = json_encode(array('error' => $error, 'res' => $res));
exit($return);
}
}
// ---------------------------- End ExcelDataStepB ---------------------------


if (isset($_POST['ExcelDataStepC'])) {
if (isset($_POST['wSeet'])) {
$error = '1';

if (!isset($_FILES["ExcelFile"])) {
$res = 'Please Upload ExcelFile!';
}
else{
$filename_icon = $_FILES["ExcelFile"]["name"];
$tempname_icon = $_FILES["ExcelFile"]["tmp_name"];
$file_basename_icon = substr($filename_icon, 0, strripos($filename_icon, '.'));
$file_ext_icon = substr($filename_icon, strripos($filename_icon, '.'));
if ($file_ext_icon != '.xlsx' && $file_ext_icon != '.xls'){
$res = 'Your file type is not supported!';
}
else if ($_POST['wSeet'] == 0) {
$res = 'Please Select WorkSheet!';
}
else if ($_POST['ExcelDataStepC'] == 0) {
$res = 'Please Select Column!';
}
else{
$error = '0';
if ($file_ext_icon == '.xlsx') {
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
}
elseif ($file_ext_icon == '.xls') {
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
}

$spreadsheet = $reader->load($tempname_icon);
$worksheet = $spreadsheet->getSheetByName($_POST['wSeet']);
$highestRow = $worksheet->getHighestRow();
$columnLetter = $_POST['ExcelDataStepC'];

$res = [];
for ($row = 2; $row <= $highestRow; $row++) {
$cellValue = $worksheet->getCell($columnLetter . $row)->getValue();
$res[$row] = $cellValue;
}

$res = json_encode($res);
}
}


$return = json_encode(array('error' => $error, 'res' => $res));
exit($return);
}
}
?>
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Ecxel </title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-T3c6CoIi6uLrA9TneNEoa7RxnatzjcDSCmG1MXxSR1GAsXEV/Dwwykc2MPK8M2HN" crossorigin="anonymous">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.7.1/jquery.min.js" integrity="sha512-v2CJ7UaYy4JwqLDIrZUI/4hqeoQieOmAZNXBeQyjo21dadnwR+8ZaIJVT8EE2iyI61OV8e6M8PP2/4hpQINQ/g==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
  </head>
  <body>
 
 <div class="p-5">

<div class="form-floating mb-3">
  <input type="file" class="form-control" id="yExcelFile" accept=".xls,.xlsx">
  <label for="yExcelFile">Upload Excel File</label>
</div>

<div class="form-floating mb-3" id="SeetSelector">
  <select class="form-select" id="wSeet"></select>
  <label for="wSeet">Select WorkSheet</label>
</div>

<div class="form-floating mb-3" id="ColumnSelector">
  <select class="form-select" id="Column"></select>
  <label for="Column">Select Column</label>
</div>


<div id="output"></div>


</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-C6RzsynM9kWDrMNeT87bh95OGNyZPhcTNXj1NW7RuBCsyN/o0jlpcV8Qyq46cDfL" crossorigin="anonymous"></script>
<script type="text/javascript">
function excelAllHide() {
$('#SeetSelector').hide();
$('#ColumnSelector').hide();
}

$(document).ready(function () { 
excelAllHide();
});
	

$('#yExcelFile').change(function () {
excelAllHide();
$('#output').html('Please Wait...');
var form_data = new FormData();
    form_data.append('ExcelDataStepA', "ExcelData");
    form_data.append('ExcelFile', $('#yExcelFile').prop('files')[0]);
$.ajax({ url: '', dataType: 'text', cache: false, contentType: false, processData: false, data: form_data, type: 'post',
success: function(result){
$('#output').hide();

data = $.parseJSON(result);
if (data.error != 1) {
$('#SeetSelector').show();
$('#wSeet').html(data.res);
}
else{
alert(data.res);
}

}
});
});



$('#wSeet').change(function () {
$('#ColumnSelector').hide();
$('#output').show();
$('#output').html('Please Wait...');
var form_data = new FormData();
    form_data.append('ExcelDataStepB', $('#wSeet').val());
    form_data.append('ExcelFile', $('#yExcelFile').prop('files')[0]);
$.ajax({ url: '', dataType: 'text', cache: false, contentType: false, processData: false, data: form_data, type: 'post',
success: function(result){
$('#output').hide();

data = $.parseJSON(result);
if (data.error != 1) {
$('#ColumnSelector').show();
$('#Column').html(data.res);
}
else{
alert(data.res);
}

}
});
});


$('#Column').change(function () {
$('#output').show();
$('#output').html('Please Wait...');
var form_data = new FormData();
    form_data.append('ExcelDataStepC', $('#Column').val());
    form_data.append('wSeet', $('#wSeet').val());
    form_data.append('ExcelFile', $('#yExcelFile').prop('files')[0]);
$.ajax({ url: '', dataType: 'text', cache: false, contentType: false, processData: false, data: form_data, type: 'post',
success: function(result){
data = $.parseJSON(result);
if (data.error != 1) {
$('#output').html(data.res);
}
else{
$('#output').hide();
alert(data.res);
}

}
});
});
</script>
</body>
</html>