<?php
//Nhúng file PHPExcel

//similar_text("Xã quang tuyến","Xã quang tuyến, Huyện đại từ, TP Thái nguyên",$percent);
//echo $percent;
//die();

require_once 'Classes/PHPExcel.php';

$file1 = 'file/file1.xlsx';
//var_dump($file1); die();
if (file_exists($file1)){
    //Tiến hành xác thực file
    $objFile1 = PHPExcel_IOFactory::identify($file1);
    $objData1 = PHPExcel_IOFactory::createReader($objFile1);

//Chỉ đọc dữ liệu
    $objData1->setReadDataOnly(true);

// Load dữ liệu sang dạng đối tượng
    $objPHPExcel1 = $objData1->load($file1);

//Lấy ra số trang sử dụng phương thức getSheetCount();
// Lấy Ra tên trang sử dụng getSheetNames();

//Chọn trang cần truy xuất
    $sheet1 = $objPHPExcel1->setActiveSheetIndex(0);

//Lấy ra số dòng cuối cùng
    $Totalrow1 = $sheet1->getHighestRow();
//Lấy ra tên cột cuối cùng
    $LastColumn1 = $sheet1->getHighestColumn();

//Chuyển đổi tên cột đó về vị trí thứ, VD: C là 3,D là 4
    $TotalCol1 = PHPExcel_Cell::columnIndexFromString($LastColumn1);

//Tạo mảng chứa dữ liệu
    $data1 = [];

//Tiến hành lặp qua từng ô dữ liệu
//----Lặp dòng, Vì dòng đầu là tiêu đề cột nên chúng ta sẽ lặp giá trị từ dòng 2
    for ($i = 0; $i <= $Totalrow1; $i++) {
        //----Lặp cột
        for ($j = 0; $j < $TotalCol1; $j++) {
            // Tiến hành lấy giá trị của từng ô đổ vào mảng
            $data1[$i][$j] = $sheet1->getCellByColumnAndRow($j, $i)->getValue();;
        }
    }
}



//doc du lieu file 2

//Đường dẫn file

$file2 = 'file/file2.xlsx';
//Tiến hành xác thực file
if (file_exists($file2)) {
    $objFile2 = PHPExcel_IOFactory::identify($file2);
    $objData2 = PHPExcel_IOFactory::createReader($objFile2);

//Chỉ đọc dữ liệu
    $objData2->setReadDataOnly(true);

// Load dữ liệu sang dạng đối tượng
    $objPHPExcel2 = $objData2->load($file2);

//Lấy ra số trang sử dụng phương thức getSheetCount();
// Lấy Ra tên trang sử dụng getSheetNames();

//Chọn trang cần truy xuất
    $sheet2 = $objPHPExcel2->setActiveSheetIndex(0);

//Lấy ra số dòng cuối cùng
    $Totalrow2 = $sheet2->getHighestRow();
//Lấy ra tên cột cuối cùng
    $LastColumn2 = $sheet2->getHighestColumn();

//Chuyển đổi tên cột đó về vị trí thứ, VD: C là 3,D là 4
    $TotalCol2 = PHPExcel_Cell::columnIndexFromString($LastColumn2);

//Tạo mảng chứa dữ liệu
    $data2 = [];

//Tiến hành lặp qua từng ô dữ liệu
//----Lặp dòng, Vì dòng đầu là tiêu đề cột nên chúng ta sẽ lặp giá trị từ dòng 2
    for ($i = 0; $i <= $Totalrow2; $i++) {
        //----Lặp cột
        for ($j = 0; $j < $TotalCol2; $j++) {
            // Tiến hành lấy giá trị của từng ô đổ vào mảng
            $data2[$i][$j] = $sheet2->getCellByColumnAndRow($j, $i)->getValue();;
        }
    }
}
//su ly du lieu
$data3 = [];
$i=0;
if (isset($_POST['xuly'])){
//    echo " --- ".$_POST['data_1_a']." --- ".$_POST['data_2_a']." --- ".$_POST['data_1_b']." --- ".$_POST['data_2_b'];
    $a1= $_POST['data_1_a'];
    $a2= $_POST['data_2_a'];
    $b1= $_POST['data_1_b'];
    $b2= $_POST['data_2_b'];

//
    foreach ($data1 as $d1){
        $val = 0;
        $value =[];
        foreach ($data2 as $d2){
            if ($d1[$a1] == $d2[$a2]){
                similar_text($d1[$b1],$d2[$b2],$percent);
                if($percent>$val){
                    $val = $percent;
                    $value = $d2;
                }
            }
        }
        $data3[$i]= array_merge($d1, $value);
        $i++;
    }
//    var_dump($data3);die();
    // luu file
    $excel = new PHPExcel();
    //Chọn trang cần ghi (là số từ 0->n)
    $excel->setActiveSheetIndex(0);
    //Tạo tiêu đề cho trang. (có thể không cần)
    $excel->getActiveSheet()->setTitle('Dữ liệu sau khi gộp');


    header("Content-Disposition: attachment; filename=\"data.xls\"");
    header("Content-Type: application/vnd.ms-excel;");
    header("Pragma: no-cache");
    header("Expires: 0");
    $out = fopen("php://output", 'w');
    foreach ($data3 as $data)
    {
        fputcsv($out, $data,"\t");
    }
    fclose($out);
    unset($_POST["xuly"]);
    return 1;

}

if (isset($_POST['postfile'])){
    if (($_FILES['file1']['name']!="")){
        if (file_exists("file/file1.xlsx")){
            unlink("file/file1.xlsx");
        }

// Where the file is going to be stored
        $target_dir = "file/";
        $file = $_FILES['file1']['name'];
        $path = pathinfo($file);
        $filename = 'file1';
        $ext = $path['extension'];
        $temp_name = $_FILES['file1']['tmp_name'];
        $path_filename_ext = $target_dir.$filename.".".$ext;

        if (file_exists($path_filename_ext)) {
            unlink("file/".$path_filename_ext);
            move_uploaded_file($temp_name,$path_filename_ext);
            echo "Tải file lên thành công!";
        }else{
            move_uploaded_file($temp_name,$path_filename_ext);
            echo "Tải file lên thành công!";
        }
        if ($ext == "xls"){
            if (file_exists("file/file1.xlsx")){
                unlink("file/file1.xlsx");
            }
            $xls_to_convert = 'file/file1.xls';
            error_reporting(E_ALL);
            ini_set('display_errors', TRUE);
            ini_set('display_startup_errors', TRUE);
            define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

            //These four lines are the entire script
//            require_once dirname(__FILE__) . '/PHPExcel/Classes/PHPExcel/IOFactory.php';
            $objPHPExcel = PHPExcel_IOFactory::load($xls_to_convert);
            $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
            $objWriter->save(str_replace('.xls', '.xlsx', $xls_to_convert));
        }
    }
    if (($_FILES['file2']['name']!="")){
        if (file_exists("file/file2.xlsx")){
            unlink("file/file2.xlsx");
        }
// Where the file is going to be stored
        $target_dir = "file/";
        $file = $_FILES['file2']['name'];
        $path = pathinfo($file);
        $filename = 'file2';
        $ext = $path['extension'];
        $temp_name = $_FILES['file2']['tmp_name'];
        $path_filename_ext = $target_dir.$filename.".".$ext;

// Check if file already exists
        if (file_exists($path_filename_ext)) {
            unlink("file/".$path_filename_ext);
            move_uploaded_file($temp_name,$path_filename_ext);
            echo "Tải file lên thành công!";
        }else{
            move_uploaded_file($temp_name,$path_filename_ext);
            echo "Tải file lên thành công";
        }
        if ($ext == "xls"){
            if (file_exists("file/file2.xlsx")){
                unlink("file/file2.xlsx");
            }
            $xls_to_convert = 'file/file2.xls';
            error_reporting(E_ALL);
            ini_set('display_errors', TRUE);
            ini_set('display_startup_errors', TRUE);
            define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

            //These four lines are the entire script
//            require_once dirname(__FILE__) . '/PHPExcel/Classes/PHPExcel/IOFactory.php';
            $objPHPExcel = PHPExcel_IOFactory::load($xls_to_convert);
            $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
            $objWriter->save(str_replace('.xls', '.xlsx', $xls_to_convert));
        }
    }
    unset($_POST["postfile"]);

}


//foreach ($data2)



//Hiển thị mảng dữ liệu
//echo '<pre>';
//var_dump($data2);


?>

<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport"
          content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Document</title>
</head>
<body style="text-align: center;">
    <form action="" method="post" style="margin: auto; margin-top: 100px;" enctype="multipart/form-data">
        <label for="file1">File 1</label>
        <input id="file1" type="file" name="file1">
        <label for="file2">File 2</label>
        <input id="file2" type="file" name="file2">
        <br>
        <input type="submit" name="postfile" value="Tải file lên" style="margin-top: 30px; background-color: #0bd398; border: none; padding: 8px;">
    </form>

    <form action="" method="post" style="margin: auto; margin-top: 100px;">
        <h4>Chọn cột để so sánh</h4>
        <table border="1" cellpadding="15" style="margin: auto">
            <tr>
                <th style="width: 200px;">Gía trị so sánh</th>
                <th style="width: 200px;">File 1</th>
                <th style="width: 200px;">File 2</th>
            </tr>
            <tr>
                <td>Tên </td>
                <td>
                    <select name="data_1_a" id="data_1_a">
                        <?php foreach ($data1[1] as $key => $data){ ?>
                        <option value="<?php echo $key ?>"><?php echo $data ?></option>
                        <?php } ?>
                    </select>
                </td>
                <td>
                    <select name="data_2_a" id="data_2_a">
                        <?php foreach ($data2[1] as $key => $data){ ?>
                            <option value="<?php echo $key ?>"><?php echo $data ?></option>
                        <?php } ?>
                    </select>
                </td>
            </tr>
            <tr>
                <td>Địa chỉ</td>
                <td>
                    <select name="data_1_b" id="data_1_b">
                        <?php foreach ($data1[1] as $key => $data){ ?>
                            <option value="<?php echo $key ?>"><?php echo $data ?></option>
                        <?php } ?>
                    </select>
                </td>
                <td>
                    <select name="data_2_b" id="data_2_b">
                        <?php foreach ($data2[1] as $key => $data){ ?>
                            <option value="<?php echo $key ?>"><?php echo $data ?></option>
                        <?php } ?>
                    </select>
                </td>
            </tr>
        </table>
        <br>
        <input type="submit" value="Xử lý ngay" name="xuly">
    </form>
</body>
</html>
