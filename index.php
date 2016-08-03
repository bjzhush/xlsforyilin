<?php
include "vendor/autoload.php";
 try {

      $file = './demo.xlsx';

      $inputFileType = PHPExcel_IOFactory::identify($file);
      $objReader = PHPExcel_IOFactory::createReader($inputFileType);

    // todo  add check code, data check
      $allSheetName = $objReader->listWorksheetNames($file);
      if ($allSheetName[0] !== 'APP') {
       exit('第一个Sheet必须是APP');
      }
      if ($allSheetName[1] !== '饿了么') {
       exit('第二个Sheet必须是饿了么');
      }
      if ($allSheetName[2] !== '美团') {
       exit('第三个Sheet必须是美团');
      }


      $objPHPExcel = $objReader->load($file);


    //获取APP数据
      $appData = [];
      $sheet0 = $objPHPExcel->getSheet(0);
      $highestRow = $sheet0->getHighestRow();
      $highestColumn = $sheet0->getHighestColumn();
      for ($row = 0; $row< $highestRow+1; $row++) {
       $isNullRow = true;
       $rowData = $sheet0->rangeToArray('A' . $row . ':' . $highestColumn . $row, NULL, TRUE, FALSE);
       $rowData = $rowData[0];
       foreach ($rowData as $cValue) {
        if(!is_null($cValue))  {
         $isNullRow = false;
        }
       }
       if($isNullRow === false) {
        $appData[]  = $rowData;
       }
      }

      $firstRow = $appData[0];
      if ($firstRow[0] !== '日期') {
       exit('APP第一列必须是日期');
      }
      if ($firstRow[1] !== '分店名称') {
       exit('APP第二列必须是分店名称');
      }

      if ($firstRow[2] !== '筛选维度') {
       exit('APP第三列必须是筛选维度');
      }

      if ($firstRow[3] !== '销售额') {
       exit('APP第四列必须是销售额');
      }

      if ($firstRow[4] !== '订单量') {
       exit('APP第五列必须是订单量');
      }

      unset($appData[0]);
      $appD = [];
      foreach ($appData as $row) {
       $appD[$row[1]] = [
           'xse' => $row[3],
           'ddl' => $row[4],
       ];
      }
      ksort($appD);

    //获取饿了么数据
      $elmData = [];
      $sheet0 = $objPHPExcel->getSheet(1);
      $highestRow = $sheet0->getHighestRow();
      $highestColumn = $sheet0->getHighestColumn();
      for ($row = 0; $row< $highestRow+1; $row++) {
       $isNullRow = true;
       $rowData = $sheet0->rangeToArray('A' . $row . ':' . $highestColumn . $row, NULL, TRUE, FALSE);
       $rowData = $rowData[0];
       foreach ($rowData as $cValue) {
        if(!is_null($cValue))  {
         $isNullRow = false;
        }
       }
       if($isNullRow === false) {
        $elmData[]  = $rowData;
       }
      }


      unset($elmData[0]);
      $elmD = [];
      foreach ($elmData as $row) {
       $elmD[$row[1]] = [
           'xse' => $row[2],
           'ddl' => $row[3],
           'dyqzk' => $row[4],
       ];
      }
      ksort($elmD);

    //获取美团数据
      $meituanData = [];
      $sheet0 = $objPHPExcel->getSheet(2);
      $highestRow = $sheet0->getHighestRow();
      $highestColumn = $sheet0->getHighestColumn();
      for ($row = 0; $row< $highestRow+1; $row++) {
       $isNullRow = true;
       $rowData = $sheet0->rangeToArray('A' . $row . ':' . $highestColumn . $row, NULL, TRUE, FALSE);
       $rowData = $rowData[0];
       foreach ($rowData as $cValue) {
        if(!is_null($cValue))  {
         $isNullRow = false;
        }
       }
       if($isNullRow === false) {
        $meituanData[]  = $rowData;
       }
      }

      unset($meituanData[0]);
      $mtD = [];
      foreach ($meituanData as $row) {
       $mtD[$row[1]] = [
           'xse' => $row[2],
           'ddl' => $row[3],
           'dyqzk' => $row[4],
       ];
      }
      ksort($mtD);

      $allDianpu = [];
      $tmp = array_keys($appD);
      foreach($tmp as $dpName) {
       $allDianpu[]  = $dpName;
      }
      $tmp = array_keys($elmD);
      foreach($tmp as $dpName) {
       $allDianpu[]  = $dpName;
      }
      $tmp = array_keys($mtD);
      foreach($tmp as $dpName) {
       $allDianpu[]  = $dpName;
      }
      $allDianpu = array_unique($allDianpu);

    //start to output csv
      $fp = fopen('result.csv', 'w');
     if ($fp === false ) {
        exit('write error');
     }
      $firstLine = '店铺,elm订单,elm销售额,elm抵用券金额,elm ASO,APP 订单,APP销售额,APP ASO,elm订单占比,elm销售额占比,elm抵用券数量,美团订单,美团销售额,美团抵用劵金额,美团ASO,美团订单占比	,美团销售额占比,美团抵用券数量';
      $firstLine = mb_convert_encoding($firstLine, 'GB2312');
      $tmpRow = explode(',', $firstLine);
      fputcsv($fp, $tmpRow);

      foreach ($allDianpu as $dianpu) {
       if (!isset($appD[$dianpu])) {
        $appD[$dianpu]  = [
            'xse' => 0,
            'ddl' => 0,
        ];
       }

       if (!isset($elmD[$dianpu])) {
        $elmD[$dianpu]  = [
            'xse' => 0,
            'ddl' => 0,
            'dyqzk' => 0,
        ];
       }

       if (!isset($mtD[$dianpu])) {
        $mtD[$dianpu]  = [
            'xse' => 0,
            'ddl' => 0,
            'dyqzk' => 0,
        ];
       }

      }


      foreach ($allDianpu as $dianpu) {
       $tmpRow = [];
       $tmpRow[] = mb_convert_encoding('U掌柜('.$dianpu.')', 'GB2312');
       $tmpRow[] = $elmD[$dianpu]['ddl'];
       $tmpRow[] = $elmD[$dianpu]['xse'];
       $tmpRow[] = $elmD[$dianpu]['dyqzk'];
       //饿了么ASO
       if ($elmD[$dianpu]['xse'] == 0) {
        $tmpRow[] = '0';
       } else {
        $tmpRow[] = number_format($elmD[$dianpu]['xse']/$elmD[$dianpu]['ddl'], 0);
       }
       $tmpRow[] = $appD[$dianpu]['ddl'];
       $tmpRow[] = $appD[$dianpu]['xse'];
       //APP ASO
       if ($appD[$dianpu]['xse'] == 0) {
        $tmpRow[] = '0';
       } else {
        $tmpRow[] = number_format($appD[$dianpu]['xse']/$appD[$dianpu]['ddl'], 0);
       }
       //饿了么订单占比
       if ($appD[$dianpu]['ddl'] == 0) {
        $tmpRow[] = '0%';
       } else {
        $tmpRow[] = number_format(100*$elmD[$dianpu]['ddl']/($elmD[$dianpu]['ddl'] + $appD[$dianpu]['ddl'] + $mtD[$dianpu]['ddl'])  ,0).'%';
       }

       //饿了么销售额占比
       if ($appD[$dianpu]['xse'] == 0) {
        $tmpRow[] = '0%';
       } else {
        $tmpRow[] = number_format(100*$elmD[$dianpu]['xse']/($elmD[$dianpu]['xse'] + $appD[$dianpu]['xse'] + $mtD[$dianpu]['xse'])  ,0).'%';
       }
       //饿了么抵用券数量
       $tmpRow[] = $elmD[$dianpu]['ddl'];

       $tmpRow[] = $mtD[$dianpu]['ddl'];
       $tmpRow[] = $mtD[$dianpu]['xse'];
       $tmpRow[] = $mtD[$dianpu]['dyqzk'];
       //美团ASO
       if ($mtD[$dianpu]['xse'] == 0) {
        $tmpRow[] = '0';
       } else {
        $tmpRow[] = number_format($mtD[$dianpu]['xse']/$mtD[$dianpu]['ddl'], 0);
       }

       //美团订单量占比
       if ($mtD[$dianpu]['ddl'] == 0) {
        $tmpRow[] = '0%';
       } else {
        $tmpRow[] = number_format(100*$mtD[$dianpu]['ddl']/($elmD[$dianpu]['ddl'] + $appD[$dianpu]['ddl'] + $mtD[$dianpu]['ddl'])  ,0).'%';
       }

       //美团销售额占比
       if ($mtD[$dianpu]['xse'] == 0) {
        $tmpRow[] = '0%';
       } else {
        $tmpRow[] = number_format(100*$mtD[$dianpu]['xse']/($elmD[$dianpu]['xse'] + $appD[$dianpu]['xse'] + $mtD[$dianpu]['xse'])  ,0).'%';
       }

       $tmpRow[] = $mtD[$dianpu]['ddl'];

       fputcsv($fp, $tmpRow);

      }

      fclose($fp);

         
 } catch (Exception $e) {
    echo "Error!";
    echo $e->getMessage(); 
    exit;
 }

?>

<a href="/result.csv">Download Result</a>
