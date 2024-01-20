<?php 
session_start();
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Write\Xlsx;
if(isset($_POST['save_excel_data']))
{
    $filename = $_FILES['import_file']['name'];
    $file_ext = pathinfo($filename, PATHINFO_EXTENSION);
    $allowed_types = ['xls','csv','xlsx'];
    
    if(in_array($file_ext , $allowed_types))
    {
        $inputFilepath = $_FILES['import_file']['tmp_name'];
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFilepath);
        $data = $spreadsheet->getActiveSheet()->toArray();
        $without_first_row=0;
        foreach($data as $row)
        {

            if($without_first_row>0)
            {
            $name = $row['0'];
            $surname = $row['1'];
            $age = $row['2'];
            $username = $row['3'];
            $passwords = $row['4'];
            try {
                $conn = new PDO("mysql:host=localhost;dbname=baza", "root","");
                $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
                $stmt = $conn->prepare("INSERT INTO `users` (`name`, `surname`, `age`, `username`, `passwords`) VALUES(:name, :surname, :age, :username, :passwords)");
                $stmt->bindParam(':name', $name);
                $stmt->bindParam(':surname', $surname);
                $stmt->bindParam(':age', $age);
                $stmt->bindParam(':username', $username);
                $stmt->bindParam(':passwords', $passwords);
                $stmt->execute();
                $bol=true;
              } 
            catch(PDOException $e) {
                echo $e->getMessage();
              } 
            }
            else {$without_first_row;}             
            $without_first_row++;        
        }
        if(isset($bol))
        {
        $_SESSION['message'] = "Imported";
        header('Location: front.php');
        exit(0);
        }
        $_SESSION['message'] = "Not Imported";
        header('Location: front.php');
        exit(0);
    }
    else
    {
        $_SESSION['message'] = "Invalid File, csv,xls,xlsx only";
        header('Location: front.php');
        exit(0);
    }
}
?>