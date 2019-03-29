<?php
/*conexion*/
//error_reporting(E_ALL);
//ini_set("display_errors", 1);
ini_set("memory_limit", "128M");
ini_set("max_execution_time", 300);
require_once '../../conexion/MySqlConexion.php';
require_once '../../conexion/configMySql.php';
/*crea obj conexion*/
$cn=MySqlConexion::getInstance();
$meses=array("","Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Setiembre","Octubre","Noviembre","Diciembre");
$az=array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
$azcount=array(5,24.5,9,9,11,11,7,9,8,6,6,6,11,4,4,4,4,4,11,11,11,6,11,10,10,10,10,10,10,10,15,15,15,15,15,15,15,15,
    15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,
    15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,
    15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,
    15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,
    15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15);
$letras=array(
    'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
// AUMENTAMOS LA CANTIDAD DE COLUMNAS
for ($i = 0; $letras[$i]; $i++) {
    for ($j = 0; $letras[$j]; $j++) {
        $az[] = $letras[$i].$letras[$j];
        $azcount[] = 15;
    }
}
$ccencap=str_replace(",","','",$_GET['ccencap']);
$nombreReporte= $_GET['nombreReporte'];
$fechainicio =$_GET['anio'] . "-" . str_pad((int)$_GET["mes"] + 1 , 2, '0',STR_PAD_LEFT) . "-" . $_GET["ini"];
$fechafin = $_GET['anio'] . "-" . str_pad((int)$_GET["mes"] + 1 , 2, '0',STR_PAD_LEFT) . "-" . $_GET["fin"];
$ayer = date("Y-m-d" , strtotime("-1 day",strtotime($fechafin)));
$anteayer = date("Y-m-d" , strtotime("-2 day",strtotime($fechafin)));
// First day of the month.
$mesPrimerDia = date('Y-m-01', strtotime($fechainicio));
// Last day of the month.
$mesUltimoDia = date('Y-m-t', strtotime($fechainicio));
$cinstit=str_replace(",","','",$_GET['cinstit']);
$diaspromedio=explode("-",$fechafin);
$sql_dias = "";
$sql_dias_column = "";
$sql_column_count = "";
$query1 = "";
$query2 = "";
$cantidadDias = $_GET["fin"] - $_GET["ini"] + 1;
for ($i = 0; $i < $cantidadDias ; $i++) {
    $cam = $i + 1;
    $dia = date("Y-m-d" , strtotime("-$i day",strtotime($fechafin)));
    $sql_dias .=" LEFT JOIN conmatp c$i ON (c$i.cconmat=c.cconmat AND c$i.fmatric='$dia')
                  LEFT JOIN gracprp g$i ON g$i.cgracpr=c$i.cgruaca ";
    $sql_dias_column .=" ,g$i.cfilial f$cam, g$i.cinstit i$cam ";
    $sql_column_count .= " ,count(g.i$cam) c$cam ";
    if ($i <= 20) {
        $query1 = "
        SELECT t.dtipcap,i.dinstit,t.ctipcap,i.cinstit
        ,count(g.it) c0
        $sql_column_count
        FROM tipcapa t
        INNER JOIN instita i
        LEFT JOIN
        (
            SELECT g.cfilial ft,i.ctipcap,g.cinstit it,c.fmatric
            $sql_dias_column
            FROM ingalum i
            INNER JOIN conmatp c ON c.cingalu=i.cingalu
            INNER JOIN gracprp g ON g.cgracpr=c.cgruaca
             INNER JOIN filialm f ON f.cfilial=g.cfilial
            $sql_dias
            WHERE c.fmatric BETWEEN '$mesPrimerDia' and '$mesUltimoDia'
            AND i.cestado=1
            AND i.ccencap IN ('$ccencap')
            GROUP BY c.cconmat
        ) g ON (t.ctipcap=g.ctipcap AND i.cinstit=g.it)
        WHERE  i.cinstit IN ('$cinstit')
        GROUP BY t.ctipcap,i.cinstit
        ORDER BY t.dtipcap,i.dinstit
    ";
        if ($i == 20) {
            // reiniciamos variables
            $sql_dias = "";
            $sql_dias_column = "";
            $sql_column_count = "";
        }
    } elseif ($i < 40) {
        $query2 = "
            SELECT t.dtipcap,i.dinstit,t.ctipcap,i.cinstit
        ,count(g.it) c0
        $sql_column_count
        FROM tipcapa t
        INNER JOIN instita i
        LEFT JOIN
        (
            SELECT g.cfilial ft,i.ctipcap,g.cinstit it,c.fmatric
            $sql_dias_column
            FROM ingalum i
            INNER JOIN conmatp c ON c.cingalu=i.cingalu
            INNER JOIN gracprp g ON g.cgracpr=c.cgruaca
             INNER JOIN filialm f ON f.cfilial=g.cfilial
            $sql_dias
            WHERE c.fmatric BETWEEN '$mesPrimerDia' and '$mesUltimoDia'
            AND i.cestado=1
            AND i.ccencap IN ('$ccencap')
            GROUP BY c.cconmat
        ) g ON (t.ctipcap=g.ctipcap AND i.cinstit=g.it)
        WHERE i.cinstit IN ('$cinstit')
        GROUP BY t.ctipcap,i.cinstit
        ORDER BY t.dtipcap,i.dinstit
        ";
    }
}
$sql = " select * from ($query1) q1 ";
if ($query2) {  $sql.= " inner join ($query2) q2 ON q2.ctipcap=q1.ctipcap AND q2.cinstit=q1.cinstit ";}
$sql .= " order BY q1.ctipcap,q1.dinstit ";
$cn->setQuery($sql);
$rpt=$cn->loadObjectList();
$sql2="SELECT concat(dnomper,' ',dappape,' ',dapmape) as nombre
        FROM personm
        WHERE dlogper='".$_GET['usuario']."'";
$cn->setQuery($sql2);
$rpt2=$cn->loadObjectList();
$sql3="SELECT dinstit
        FROM instita
        WHERE cinstit IN ('$cinstit')
        ORDER BY dinstit";
$cn->setQuery($sql3);
$rpt3=$cn->loadObjectList();
/*
echo count($control)."-";
echo $sql;
*/
date_default_timezone_set('America/Lima');
require_once 'includes/Classes/PHPExcel.php';
$styleThinBlackBorderOutline = array(
    'borders' => array(
        'outline' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('argb' => 'FF000000'),
        ),
    ),
);
$styleThinBlackBorderAllborders = array(
    'borders' => array(
        'allborders' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('argb' => 'FF000000'),
        ),
    ),
);
$styleThickBlackBorderOutline = array(
    'borders' => array(
        'outline' => array(
            'style' => PHPExcel_Style_Border::BORDER_THICK,
            'color' => array('argb' => 'FF000000'),
        ),
    ),
);
$styleThickBlackBorderRight = array(
    'borders' => array(
        'right' => array(
            'style' => PHPExcel_Style_Border::BORDER_THICK,
            'color' => array('argb' => 'FF000000'),
        ),
    ),
);
$styleThickBlackBorderAllborders = array(
    'borders' => array(
        'allborders' => array(
            'style' => PHPExcel_Style_Border::BORDER_THICK,
            'color' => array('argb' => 'FF000000'),
        ),
    ),
);
$styleBold= array(
    'font'    => array(
        'bold'      => true
    )
);
$styleAlignmentBold= array(
    'font'    => array(
        'bold'      => true
    ),
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
    ),
);
$styleAlignment= array(
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
    ),
);
$styleAlignmentRight= array(
    'font'    => array(
        'bold'      => true
    ),
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
        'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
    ),
);
$styleColor = array(
    'fill' => array(
        'type' => PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
        'startcolor' => array(
            'argb' => 'FFA0A0A0',
        ),
        'endcolor' => array(
            'argb' => 'FFFFFFFF',
        )
    ),
);
function color(){
    $color=array(0,1,2,3,4,5,6,7,8,9,"A","B","C","D","E","F");
    $dcolor="";
    for($i=0;$i<6;$i++){
        $dcolor.=$color[rand(0,15)];
    }
    $num='FA'.$dcolor;
    $styleColorFunction = array(
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
            'startcolor' => array(
                'argb' => $num,
            ),
            'endcolor' => array(
                'argb' => 'FFFFFFFF',
            )
        ),
    );
    return $styleColorFunction;
}
$colorVerde = "FFEBF1DE";
$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setCreator("Jorge Salcedo")
    ->setLastModifiedBy("Jorge Salcedo")
    ->setTitle("Office 2007 XLSX Test Document")
    ->setSubject("Office 2007 XLSX Test Document")
    ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
    ->setKeywords("office 2007 openxml php")
    ->setCategory("Test result file");
$objPHPExcel->getDefaultStyle()->getFont()->setName('Bookman Old Style');
$objPHPExcel->getDefaultStyle()->getFont()->setSize(8);
$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
$objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
$objPHPExcel->getActiveSheet()->setCellValue("A1","REPORTE DE MATRÍCULA MEDIOS GENERALES - ".$nombreReporte);
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setSize(20);
$objPHPExcel->getActiveSheet()->setCellValue("B2","MES: " . strtoupper($meses[$_GET["mes"] + 1]));
$objPHPExcel->getActiveSheet()->getStyle('B2')->getFont()->setSize(12);
$objPHPExcel->getActiveSheet()->getStyle('B2')->applyFromArray($styleAlignmentBold);
$cabecera=array('N°','MEDIO');
$cantidadaz=1;
for($i=1;$i<= $cantidadDias * 1 + 2;$i++){
    $cantidadaz++;
    //$azcount[$cantidadaz]=5.5;
    array_push($cabecera, 'TOTALES');
    IF($i == 1)
    {   $azcount[$cantidadaz]=4.1;  }
    IF($i == 2)
    {   $azcount[$cantidadaz]=3.9;  }
    IF($i > 2)
    {   $azcount[$cantidadaz]=3.4;  }
    foreach($rpt3 as $r){
        $cantidadaz++;
        //$azcount[$cantidadaz]=5.5;
        IF($i == 1)
        {   $azcount[$cantidadaz]=4.1;  }
        IF($i == 2)
        {   $azcount[$cantidadaz]=3.9;  }
        IF($i > 2)
        {   $azcount[$cantidadaz]=3.4;  }
        array_push($cabecera, $r['dinstit']);
    }
}
for($i=0;$i<count($cabecera);$i++){
    $objPHPExcel->getActiveSheet()->setCellValue($az[$i]."6",$cabecera[$i]);
    if($i>=2){
        $objPHPExcel->getActiveSheet()->getStyle($az[$i]."6")->getAlignment()->setTextRotation(90);
    }
    $objPHPExcel->getActiveSheet()->getStyle($az[$i]."6")->getAlignment()->setWrapText(true);
    $objPHPExcel->getActiveSheet()->getColumnDimension($az[$i])->setWidth($azcount[$i]);
}
$objPHPExcel->getActiveSheet()->mergeCells('A1:'.$az[($i-1)].'1');
$objPHPExcel->getActiveSheet()->getStyle('A1:'.$az[($i-1)].'1')->applyFromArray($styleAlignmentBold);
$objPHPExcel->getActiveSheet()->getRowDimension("5")->setRowHeight(46.5); // altura
$objPHPExcel->getActiveSheet()->getStyle('A5:'.$az[($i-1)].'6')->applyFromArray($styleAlignmentBold);
$objPHPExcel->getActiveSheet()->mergeCells('B3:C3');
$objPHPExcel->getActiveSheet()->setCellValue("B3","USUARIO:");
$objPHPExcel->getActiveSheet()->getStyle('B3:C3')->applyFromArray($styleAlignmentRight);
$objPHPExcel->getActiveSheet()->setCellValue("D3",$rpt2[0]['nombre']);
$objPHPExcel->getActiveSheet()->mergeCells('B4:C4');
$objPHPExcel->getActiveSheet()->setCellValue("B4","FECHA IMPRESIÓN");
$objPHPExcel->getActiveSheet()->getStyle('B4:C4')->applyFromArray($styleAlignmentRight);
$objPHPExcel->getActiveSheet()->setCellValue("D4",date("Y-m-d"));
$objPHPExcel->getActiveSheet()->setCellValue("B5","FECHA MATRÍCULA \n".$mesPrimerDia ."/" .$mesUltimoDia);
$objPHPExcel->getActiveSheet()->getStyle("B5")->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->setCellValue("C5", 'PROMEDIO');
$objPHPExcel->getActiveSheet()->mergeCells('C5:'.$az[(2+count($rpt3)*1+1-1)]."5");
$objPHPExcel->getActiveSheet()->setCellValue($az[(2+count($rpt3)*1+1)]."5", 'CONSOLIDADO');
$objPHPExcel->getActiveSheet()->mergeCells($az[(2+count($rpt3)*1+1)]."5:".$az[(2+count($rpt3)*2+2-1)]."5");
// GENERA LAS CABECERAS DE LAS FECHAS
for ($i = 0; $i < $cantidadDias ; $i++) {
    $dia = date("Y-m-d" , strtotime("-$i day",strtotime($fechafin)));
    $x = $i + 2;
    $y = $x + 1;
    $objPHPExcel->getActiveSheet()->setCellValue($az[(2+count($rpt3)*$x+$x)]."5", $dia);
    $objPHPExcel->getActiveSheet()->mergeCells($az[(2+count($rpt3)*$x+$x)]."5:".$az[(2+count($rpt3)*$y+$y-1)]."5");
}
$objPHPExcel->getActiveSheet()->getStyle("B5:".$az[(2+count($rpt3)*$y+$y-1)]."5")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFEBF1DE');
$objPHPExcel->getActiveSheet()->getStyle("A6:".$az[(2+count($rpt3)*$y+$y-1)]."6")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFEBF1DE');
/**/
$valorinicial=6; // filas
$cont=0;
$countrpt3=0;
$arrayFilas = array();
foreach($rpt as $r){
    $posicionaz=0;
    $countrpt3++;
    if ($countrpt3 == 1) {
        $cont++;
        $valorinicial="{fila}"; // usamos un comodin para luego poner la fila correspondie debido al reordenamiento que se va a dar
        $arrayFilas[$r['dtipcap']][$az[$posicionaz]] = $valorinicial;
        $arrayFilas[$r['dtipcap']][$az[++$posicionaz]] = $r['dtipcap'];
        $arrayFilas[$r['dtipcap']][$az[++$posicionaz]] = '=SUM(D'.$valorinicial.':'.$az[(2+count($rpt3)*1+1-1)].$valorinicial.")";
        for ($i = 1; $i <= $cantidadDias + 1 ; $i++) {  // totales por grupo  desde consolidado
            $x = $i + 1;
            $arrayFilas[$r['dtipcap']][$az[(2+count($rpt3)*$i+$i)]] =  "=SUM(".$az[(2+count($rpt3)*$i+$i+1)].$valorinicial.":".$az[(2+count($rpt3)*$x+$x-1)].$valorinicial.")";
        }
        for($i=1;$i<=count($rpt3);$i++){
            //promedios por institucion , solo para el primer grupo (promedios)
            $arrayFilas[$r['dtipcap']][$az[(2+$i)]] = '='.$az[(2+$i+count($rpt3)+1)].$valorinicial.'/'.($diaspromedio[2]-1);
        }
    }
    $posicionaz=2+count($rpt3)+$countrpt3;
    $posicionaz++;
    $arrayFilas[$r['dtipcap']][$az[$posicionaz]] = $r['c0'];
    $arrayFilas[$r['dtipcap']]["consolidadototal"] += $r['c0'];
    $posicionaz += count($rpt3);
    for ($i = 1; $i <= $cantidadDias; $i++) { // llena la misma institucion por cadad grupo
        $posicionaz++;
        $arrayFilas[$r['dtipcap']][$az[$posicionaz]] = $r['c'.$i];
        $posicionaz = ($i == $cantidadDias) ? $posicionaz++ : $posicionaz+=count($rpt3);
    }
    if( $countrpt3==count($rpt3) ){
        $countrpt3=0;
    }
}
// ordernarlos por consolidado total
foreach ($arrayFilas as $key => $row) {
    $rows[$key]  = $row["consolidadototal"];
}
array_multisort($rows, SORT_DESC, $arrayFilas);
// agregar filas la excel
$valorinicial=6; // filas
foreach ($arrayFilas as $k1 => $v1) {
    if ((int)$v1["consolidadototal"]) {
        $valorinicial++;
        foreach ($v1 as $columna => $valor) {
            if ($columna != "consolidadototal"){
                $objPHPExcel->getActiveSheet()->setCellValue($columna.$valorinicial, str_replace("{fila}", $valorinicial, $valor));
                if ($valor > 0 || strpos($valor, "=SUM(") === true) {
                    $objPHPExcel->getActiveSheet()->getStyle($columna.$valorinicial)->applyFromArray($styleAlignmentBold);
                }
            }
        }
    }
}
/**/
///negrita a los numeros
$objPHPExcel->getActiveSheet()->getStyle("C6:".$az[(2+count($rpt3)*2+2-1)].$valorinicial)->applyFromArray($styleBold);
$objPHPExcel->getActiveSheet()->getStyle($az[(2+count($rpt3)*2+2)]."6:".$az[(2+count($rpt3)*2+2)].$valorinicial)->applyFromArray($styleBold);
$objPHPExcel->getActiveSheet()->getStyle($az[(2+count($rpt3)*3+3)]."6:".$az[(2+count($rpt3)*3+3)].$valorinicial)->applyFromArray($styleBold);
$objPHPExcel->getActiveSheet()->getStyle('B5:'.$az[$cantidadaz]."5")->applyFromArray($styleThinBlackBorderAllborders);
$objPHPExcel->getActiveSheet()->getStyle('A6:'.$az[$cantidadaz].$valorinicial)->applyFromArray($styleThinBlackBorderAllborders);
$valorinicial++;
$cantidadaz=1;
$objPHPExcel->getActiveSheet()->setCellValue($az[$cantidadaz].$valorinicial, "TOTALES");
for($i=1;$i<=$cantidadDias + 2;$i++){
    $cantidadaz++;
    $objPHPExcel->getActiveSheet()->setCellValue($az[$cantidadaz].$valorinicial, "=SUM(".$az[$cantidadaz]."7:".$az[$cantidadaz].($valorinicial-1).")");
    for($j=0;$j<count($rpt3);$j++){
        $cantidadaz++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$cantidadaz].$valorinicial, "=SUM(".$az[$cantidadaz]."7:".$az[$cantidadaz].($valorinicial-1).")");
    }
}
// PINTADODE LAS FILA TOTALES
$objPHPExcel->getActiveSheet()->getStyle("B$valorinicial:".$az[$cantidadaz].$valorinicial)
    ->getFill()
    ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
    ->getStartColor()
    ->setARGB('FFEBF1DE')
;
$objPHPExcel->getActiveSheet()->getStyle('B'.$valorinicial.":".$az[$cantidadaz].$valorinicial)->applyFromArray($styleThinBlackBorderAllborders);
$objPHPExcel->getActiveSheet()->getStyle('B'.$valorinicial.":".$az[$cantidadaz].$valorinicial)->applyFromArray($styleBold);
// PINTANDO MARGENES
$lastColumn = $az[count($cabecera)-1];
$lastrow = $valorinicial;
$cantInst = count($rpt3);
$objPHPExcel->getActiveSheet()->getStyle("B5:".$lastColumn.(5))->applyFromArray($styleThickBlackBorderAllborders);
$objPHPExcel->getActiveSheet()->getStyle("A6:B6")->applyFromArray($styleThickBlackBorderAllborders);
$objPHPExcel->getActiveSheet()->getStyle("A7:A".($lastrow*1 - 1) )->applyFromArray($styleThickBlackBorderOutline);
$objPHPExcel->getActiveSheet()->getStyle("B7:B".($lastrow - 1) )->applyFromArray($styleThickBlackBorderOutline);
// pintar grupos
$grupoColumnInicial = 2;
for ($i = 0; $i < $cantidadDias + 2;$i++){
    $objPHPExcel->getActiveSheet()->getStyle($az[$grupoColumnInicial]."6:".$az[$grupoColumnInicial + count($rpt3)]."6")->applyFromArray($styleThickBlackBorderOutline);
    $objPHPExcel->getActiveSheet()->getStyle($az[$grupoColumnInicial]."7:".$az[$grupoColumnInicial].($lastrow - 1))->applyFromArray($styleThickBlackBorderOutline);
    $grupoColumnInicial = count($rpt3) + $grupoColumnInicial + 1;
}
$objPHPExcel->getActiveSheet()->getStyle("$lastColumn".(7).":$lastColumn$lastrow")->applyFromArray($styleThickBlackBorderRight);
$objPHPExcel->getActiveSheet()->getStyle("B".$lastrow.":$lastColumn$lastrow")->applyFromArray($styleThickBlackBorderAllborders);
////////////////////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->setTitle('Medio_Generales_Matricula');
// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);
// Redirect output to a client's web browser (Excel2007)
/*header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="Medio_Generales_Matricula_'.date("Y-m-d_H-i-s"). " - ".$nombreReporte .'.xlsx"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');*/
header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="Medio_Generales_Matricula_'.date("Y-m-d_H-i-s").'.xls"'); // file name of excel
            header('Cache-Control: max-age=0');
            // If you're serving to IE 9, then the following may be needed
            header('Cache-Control: max-age=1');
            // If you're serving to IE over SSL, then the following may be needed
            header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
            header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
            header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
            header ('Pragma: public'); // HTTP/1.0
            $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
            $objWriter->save('php://output');
exit;
?>
