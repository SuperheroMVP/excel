
<?php
//Nhúng file PHPExcel

//similar_text("Xã quang tuyến","Xã quang tuyến, Huyện đại từ, TP Thái nguyên",$percent);
//echo $percent;
//die();

$con=mysqli_connect('localhost','root','','excel');

require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';
//require_once 'PHPExcel/Classes/PHPExcel/IOFactory.php';
$data1 = [];
$data2 =[];
$data3 = [];
$data4 =[];
if (isset($_POST['postfile'])){
    $allowedFileType = ['application/vnd.ms-excel','text/xls','text/xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];

      if (($_FILES['file1']['name']!=""))     
      {
        if(in_array($_FILES["file1"]["type"],$allowedFileType))
        {
           $file1 = $_FILES['file1']['tmp_name'];
           $objExcel1 = PHPExcel_IOFactory::load($file1);

           foreach ($objExcel1->getWorksheetIterator() as $worksheet) 
           {

               $highestrow = $worksheet->getHighestRow();
               $highestcolumn = $worksheet->getHighestColumn();

               for ($i = 0; $i <= $highestrow; $i++) 
               {
                   $name = $worksheet->getCellByColumnAndRow(0, $i)->getValue();
                   $address = $worksheet->getCellByColumnAndRow(1, $i)->getValue();
                   $phone = $worksheet->getCellByColumnAndRow(2, $i)->getValue();

                   // if($name!='')
                   // {
                   //     $insertqry="INSERT INTO `tbl_excel1`( `name`, `address`,`phone`) VALUES ('$name','$address','$phone')";
                   //     $insertres = mysqli_query($con,$insertqry);
                   // }
                   $data1 = array("$name", "$address", "$phone");
                   array_push($data2, $data1);
                }

                echo "upload file 1 success .</br>";
            }
        }
        else
        {
          echo "File 1:Incorrect file format. Upload Excel File.</br>";
        }
      }
      else
      {
          
          echo "File 1: You not choose file.</br>";
      }

      if (($_FILES['file2']['name']!=""))     
      {
        if(in_array($_FILES["file2"]["type"],$allowedFileType))
        {
           $file2 = $_FILES['file2']['tmp_name'];
           $objExcel2 = PHPExcel_IOFactory::load($file2);

           foreach ($objExcel2->getWorksheetIterator() as $worksheet) 
           {

               $highestrow = $worksheet->getHighestRow();
               $highestcolumn = $worksheet->getHighestColumn();

               for ($i = 0; $i <= $highestrow; $i++) 
               {
                   $name = $worksheet->getCellByColumnAndRow(0, $i)->getValue();
                   $address = $worksheet->getCellByColumnAndRow(1, $i)->getValue();
                   $phone = $worksheet->getCellByColumnAndRow(2, $i)->getValue();
                   $job = $worksheet->getCellByColumnAndRow(3, $i)->getValue();
                   // if($name!='')
                   // {
                   //     $insertqry2="INSERT INTO `tbl_excel2`( `name`,`address`,`phone`,`job`) VALUES ('$name','$address','$phone','$job')";
                   //     $insertres = mysqli_query($con,$insertqry2);
                   // }
                   $data3 = array("$name", "$address","$phone","$job");
                   array_push($data4, $data3);
                }

                echo "upload file 2 success .</br>";
            }
        }
        else
        {
          echo "File 2:Incorrect file format. Upload Excel File.</br>";
        }
      }
      else
      {
          
          echo "File 2: You not choose file.</br>";
      }

  for($f1=0;$f1<count($data2);$f1++)
  {
    for($f2=0;$f2< count($data4);$f2++)
    {
       if(similar_text( $f1[1], $f2[1], $percent))
       {
         echo "trung.</br>";
       }
    }
  }
   
}


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

<form method="post" style="margin: auto; margin-top: 100px;">
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

                    <?php
                        foreach ($data2[1] as $key => $data){
                            ?>
                               <option value="<?php echo $key ?>"><?php echo $data ?></option>
                           <?php
                        }
                    ?>
                </select>
            </td>
            <td>
                <select name="data_2_a" id="data_2_a">
                    <?php
                        foreach ($data4[1] as $key => $data){
                            ?>
                               <option value="<?php echo $key ?>"><?php echo $data ?></option>
                           <?php
                        }
                    ?>
                </select>
            </td>
        </tr>
        <tr>
            <td>Địa chỉ</td>
            <td>
                <select name="data_1_b" id="data_1_b">
                    <?php
                        foreach ($data2[1] as $key => $data){
                            ?>
                               <option value="<?php echo $key ?>"><?php echo $data ?></option>
                           <?php
                        }
                    ?>
                </select>
            </td>
            <td>
                <select name="data_2_b" id="data_2_b">
                    <?php
                        foreach ($data4[1] as $key => $data){
                            ?>
                               <option value="<?php echo $key ?>"><?php echo $data ?></option>
                           <?php
                        }
                    ?>
                </select>
            </td>
        </tr>
    </table>
    <br>
    <input type="submit" value="Xử lý ngay" name="xuly">
</form>
</body>
</html>
