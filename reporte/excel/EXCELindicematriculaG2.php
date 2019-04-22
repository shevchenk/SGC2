<?php
/*conexion*/
require_once '../../conexion/MySqlConexion.php';
require_once '../../conexion/configMySql.php';

/*crea obj conexion*/
$cn=MySqlConexion::getInstance();

$az=array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ','CA','CB','CC','CD','CE','CF','CG','CH','CI','CJ','CK','CL','CM','CN','CO','CP','CQ','CR','CS','CT','CU','CV','CW','CX','CY','CZ','DA','DB','DC','DD','DE','DF','DG','DH','DI','DJ','DK','DL','DM','DN','DO','DP','DQ','DR','DS','DT','DU','DV','DW','DX','DY','DZ');
$azcount=array(2.5,9.2,10.2,26,7.2,11.6,0,0,8.5,
                2.7,2.7,2.7,2.7,2.7,2.7,2.7,2.7,2.7,2.7,2.7,2.7,2.7,2.7,
                5.3,6.6,5.6,4.7,4.6,8,4.6,5.3,5,5.6,6.2,6.2,5,10,
            15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,);

$fechas=" '".$_GET['fechini']."' AND '".$_GET['fechfin']."'";
$cfilial=str_replace(",","','",$_GET['cfilial']);
$cinstit=str_replace(",","','",$_GET['cinstit']);

//despues de dias_falta  ,MAX(CONCAT(cc.cestado,IF(cc.ccarrer='','000',cc.ccarrer),cc.fusuari,'|',cc.nprecio)) precio 
//despues de gracprp  INNER JOIN cropaga cr ON (cr.cgruaca=g.cgracpr AND cr.ccuota='2' AND cr.cestado='1')
//despues de cropaga  INNER JOIN concepp cc ON (cc.cconcep=cr.cconcep AND (cc.ccarrer in ('',g.ccarrer)) )
$sql="  SELECT f.dfilial,i.dinstit,c.dcarrer
        ,(  SELECT GROUP_CONCAT(d.dnemdia SEPARATOR '-') 
            FROM diasm d 
            WHERE FIND_IN_SET (d.cdia,replace(g.cfrecue,'-',','))  >  0
        ) AS frec
        ,concat(h.hinici,' - ',h.hfin) AS hora
        ,g.csemaca,g.cinicio,DATE_FORMAT(g.finicio, '%d/%m/%y') AS finicio
        ,COUNT(DISTINCT(ing.cingalu)) AS inscritos,g.nmetmat,g.nmetmin
        ,COUNT(DISTINCT( IF(co.fmatric=(CURDATE()-interval 13 day) AND ing.cestado='1',ing.cingalu,NULL) )) AS d1
        ,COUNT(DISTINCT( IF(co.fmatric=(CURDATE()-interval 12 day) AND ing.cestado='1',ing.cingalu,NULL) )) AS d2
        ,COUNT(DISTINCT( IF(co.fmatric=(CURDATE()-interval 11 day) AND ing.cestado='1',ing.cingalu,NULL) )) AS d3
        ,COUNT(DISTINCT( IF(co.fmatric=(CURDATE()-interval 10 day) AND ing.cestado='1',ing.cingalu,NULL) )) AS d4
        ,COUNT(DISTINCT( IF(co.fmatric=(CURDATE()-interval 9 day) AND ing.cestado='1',ing.cingalu,NULL) )) AS d5
        ,COUNT(DISTINCT( IF(co.fmatric=(CURDATE()-interval 8 day) AND ing.cestado='1',ing.cingalu,NULL) )) AS d6
        ,COUNT(DISTINCT( IF(co.fmatric=(CURDATE()-interval 7 day) AND ing.cestado='1',ing.cingalu,NULL) )) AS d7
        ,COUNT(DISTINCT( IF(co.fmatric=(CURDATE()-interval 6 day) AND ing.cestado='1',ing.cingalu,NULL) )) AS d8
        ,COUNT(DISTINCT( IF(co.fmatric=(CURDATE()-interval 5 day) AND ing.cestado='1',ing.cingalu,NULL) )) AS d9
        ,COUNT(DISTINCT( IF(co.fmatric=(CURDATE()-interval 4 day) AND ing.cestado='1',ing.cingalu,NULL) )) AS d10
        ,COUNT(DISTINCT( IF(co.fmatric=(CURDATE()-interval 3 day) AND ing.cestado='1',ing.cingalu,NULL) )) AS d11
        ,COUNT(DISTINCT( IF(co.fmatric=(CURDATE()-interval 2 day) AND ing.cestado='1',ing.cingalu,NULL) )) AS d12
        ,COUNT(DISTINCT( IF(co.fmatric=(CURDATE()-interval 1 day) AND ing.cestado='1',ing.cingalu,NULL) )) AS d13
        ,COUNT(DISTINCT( IF(co.fmatric=CURDATE() AND ing.cestado='1' ,ing.cingalu,NULL) )) AS d14
        ,DATE_FORMAT(s.finimat, '%d/%m/%y') AS inicamp
        ,DateDiff(curdate(),s.finimat) AS ndiacamp
        ,IF(DateDiff(curdate(),g.finicio) >=0,0,(DateDiff(g.finicio,curdate()) )) AS dias_falta
        , g.observacion
        FROM gracprp g
        INNER JOIN filialm f ON f.cfilial=g.cfilial
        INNER JOIN instita i ON i.cinstit=g.cinstit
        INNER JOIN carrerm c ON c.ccarrer=g.ccarrer
        INNER JOIN semacan s ON ( s.cfilial=g.cfilial AND s.cinstit=g.cinstit AND s.csemaca=g.csemaca AND s.cinicio=g.cinicio )
        INNER JOIN horam h ON h.chora=g.chora
        LEFT JOIN conmatp co ON co.cgruaca=g.cgracpr
        LEFT JOIN ingalum ing ON (ing.cingalu=co.cingalu AND ing.cestado='1')
        WHERE g.cesgrpr IN ('3','4')
        AND g.finicio BETWEEN $fechas
        AND g.cfilial IN ('$cfilial')
        AND g.cinstit IN ('$cinstit')
        GROUP BY g.csemaca,g.cfilial,g.cinstit,g.ccarrer,g.cciclo,
        g.cinicio,g.cfrecue,g.cturno,g.chora,g.finicio
        ORDER BY f.dfilial,i.dinstit,c.dcarrer,frec";
$cn->setQuery($sql);
$rpt=$cn->loadObjectList();


$sql2="Select concat(dnomper,' ',dappape,' ',dapmape) as nombre
        FROM personm
        WHERE dlogper='".$_GET['usuario']."'";
$cn->setQuery($sql2);
$rpt2=$cn->loadObjectList();

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
$styleAlignmentLeft= array(
    'alignment' => array(
        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
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

$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setCreator("Jorge Salcedo")
                             ->setLastModifiedBy("Jorge Salcedo")
                             ->setTitle("Indice de Matriculacion para gerencia")
                             ->setSubject("Reporte IMG")
                             ->setDescription("Agiliza la toma de desiciones en los indice de matriculación realizadas según el filtro realizado.")
                             ->setKeywords("php")
                             ->setCategory("Reporte");

$objPHPExcel->getDefaultStyle()->getFont()->setName('Bookman Old Style');
$objPHPExcel->getDefaultStyle()->getFont()->setSize(8);

$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
$objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
$objPHPExcel->getActiveSheet()->getPageSetup()->setFitToPage(false);
$objPHPExcel->getActiveSheet()->getPageSetup()->setScale(89);
$objPHPExcel->getActiveSheet()->getPageMargins()->setTop(0.7);
$objPHPExcel->getActiveSheet()->getPageMargins()->setRight(0.2);
$objPHPExcel->getActiveSheet()->getPageMargins()->setLeft(0.75);
$objPHPExcel->getActiveSheet()->getPageMargins()->setBottom(0.1);
$objPHPExcel->getActiveSheet()->getPageMargins()->setHeader(0.20);
$objPHPExcel->getActiveSheet()->getPageMargins()->setFooter(0);


//$objPHPExcel->getActiveSheet()->setCellValue("A1",$sql);
$objPHPExcel->getActiveSheet()->setCellValue("A2","ÍNDICE DE MATRICULACIÓN PARA GERENCIA");
$objPHPExcel->getActiveSheet()->getStyle('A2')->getFont()->setSize(20);
$objPHPExcel->getActiveSheet()->mergeCells('A2:AI2');
$objPHPExcel->getActiveSheet()->getStyle('A2:AI2')->applyFromArray($styleAlignmentBold);

$cabeceraG=array('N°','ODE','INSTITUCION','CARRERA','FREC','HORA','SEMESTRE','INICIO','FECHA INICIO','MATRÍCULAS SEMANA ANTERIOR','','','','','','','MATRÍCULAS ÚLTIMOS 7 DÍAS','','','','','','','SEMANA ANTERIOR','ÚLTIMOS 7 DÍAS','TOTAL INSCRITOS','META MÁXIMA','META MÍNIMA','INCIO CAMPAÑA','DIAS CAMPAÑA','DÍAS QUE FALTA','INDICE DIARIO','','PROY. DIAS FALTANTES','PROY. FINAL','FALTA PARA LOGRAR META', 'OBSERVACION');

    for($i=0;$i<count($cabeceraG);$i++){
    $objPHPExcel->getActiveSheet()->setCellValue($az[$i]."4",$cabeceraG[$i]);
        if( ($i>=23 AND $i<=30) OR ($i>=33 AND $i<=35) ){
            $objPHPExcel->getActiveSheet()->getStyle($az[$i]."4")->getAlignment()->setTextRotation(90);
        }
        if( ($i>=0 AND $i<=8) OR ($i>=23 AND $i<=30) OR ($i>=33 AND $i<=36) ){
            $objPHPExcel->getActiveSheet()->mergeCells($az[$i]."4:".$az[$i]."5");
        }
    $objPHPExcel->getActiveSheet()->getStyle($az[$i]."4")->getAlignment()->setWrapText(true);
    $objPHPExcel->getActiveSheet()->getColumnDimension($az[$i])->setWidth($azcount[$i]);
    }
    $objPHPExcel->getActiveSheet()->mergeCells('J4:P4');
    $objPHPExcel->getActiveSheet()->mergeCells('Q4:W4');
    $objPHPExcel->getActiveSheet()->mergeCells('AF4:AG4');
    $objPHPExcel->getActiveSheet()->getRowDimension("4")->setRowHeight(23.75); // altura
    $objPHPExcel->getActiveSheet()->getRowDimension("5")->setRowHeight(53.75); // altura


$i=3;
$objPHPExcel->getActiveSheet()->mergeCells('B'.$i.':C'.$i);
$objPHPExcel->getActiveSheet()->setCellValue("B".$i,"FECHA IMPRESIÓN");
$objPHPExcel->getActiveSheet()->getStyle('B'.$i.':C'.$i)->applyFromArray($styleAlignmentRight);
$objPHPExcel->getActiveSheet()->setCellValue("D".$i,date("Y-m-d"));

$i++;
$objPHPExcel->getActiveSheet()->getStyle('A'.$i.':AK5' )->applyFromArray($styleAlignmentBold);
$objPHPExcel->getActiveSheet()->getStyle('A'.$i.':AK5' )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFEBF1DE');

$valorinicial=5;
$objPHPExcel->getActiveSheet()->setCellValue("AF".$valorinicial,"SEMANA ANTERIOR");
$objPHPExcel->getActiveSheet()->setCellValue("AG".$valorinicial,"ÚLTIMOS 7 DÍAS");
$objPHPExcel->getActiveSheet()->getStyle("AF5:AG5")->getAlignment()->setTextRotation(90);
$objPHPExcel->getActiveSheet()->getStyle("AF5:AG5")->getAlignment()->setWrapText(true);

$fechafin=date('y-m-d');
$fechaini=date('y-m-d',strtotime('-13 days'));
$paz=9;
while ( $fechaini<= $fechafin) {
    $objPHPExcel->getActiveSheet()->setCellValue($az[$paz]."5",$fechaini); $paz++;
    $fechaini=date('y-m-d',strtotime('+1 day',strtotime($fechaini)));
}
    $objPHPExcel->getActiveSheet()->getStyle("J5:W5")->getAlignment()->setTextRotation(90);
    $objPHPExcel->getActiveSheet()->getStyle("J5:W5")->getAlignment()->setWrapText(true);
//$objPHPExcel->getActiveSheet()->getStyle("Q5:V5")->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0F000'); 

$cont=0;
$total=0;
$pago="";
$fil="";
$cins="";
$frec="";
$hora="";
$concarrera=0;
$dcarrer='';
$sumatotal=array();

foreach($rpt as $r){    
    if($fil!=$r['dfilial'] /*or $cins!=$r['dinstit']*/){
        if($fil!=''){   
            /*if($dcarrer!=''){
            $objPHPExcel->getActiveSheet()->mergeCells( "D".($valorinicial-$concarrera).":D".($valorinicial) );
            $objPHPExcel->getActiveSheet()->getStyle("D".($valorinicial-$concarrera).":D".($valorinicial))->applyFromArray($styleAlignmentLeft);
            }*/     
        $valorinicial++;            
        $objPHPExcel->getActiveSheet()->getRowDimension($valorinicial)->setRowHeight(15.5); // altura
        $objPHPExcel->getActiveSheet()->setCellValue("I".$valorinicial,"TOTALES");
        $objPHPExcel->getActiveSheet()->getStyle('I'.$valorinicial.":"."AK".$valorinicial)->applyFromArray($styleAlignmentRight);
        $objPHPExcel->getActiveSheet()->getStyle("A".$valorinicial.":"."AK".$valorinicial)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFD8E4BC');
        
        $col=9;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++; $col++; $col++; $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;

        array_push($sumatotal, $valorinicial);
        /*$objPHPExcel->getActiveSheet()->setCellValue( "B".($valorinicial-$cont), $fil );
        $objPHPExcel->getActiveSheet()->mergeCells( "B".($valorinicial-$cont).":B".($valorinicial-1) );
        $objPHPExcel->getActiveSheet()->setCellValue( "C".($valorinicial-$cont), $cins );
        $objPHPExcel->getActiveSheet()->mergeCells( "C".($valorinicial-$cont).":C".($valorinicial-1) );
        $objPHPExcel->getActiveSheet()->getStyle( "B".($valorinicial-$cont).":C".($valorinicial-1) )->applyFromArray($styleAlignment);
        */
        }

    $fil=$r['dfilial'];
    /*$cins=$r['dinstit'];*/
    $cont=0;
    /*$concarrera=0;
    $dcarrer='';*/
    }
$cont++;
//$concarrera++;
$valorinicial++;
$paz=0;

$objPHPExcel->getActiveSheet()->getRowDimension($valorinicial)->setRowHeight(25.25); // altura
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $cont);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['dfilial']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['dinstit']);$paz++;

/*if($r["dcarrer"]!=$dcarrer){
    if($dcarrer!=''){
        $objPHPExcel->getActiveSheet()->mergeCells( "D".($valorinicial-$concarrera).":D".($valorinicial-1) );
        $objPHPExcel->getActiveSheet()->getStyle("D".($valorinicial-$concarrera).":D".($valorinicial))->applyFromArray($styleAlignmentLeft);
    }
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['dcarrer']);

$concarrera=0;
$dcarrer=$r["dcarrer"];
}*/
//$precio=explode("|",$r['precio']);
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['dcarrer']);$paz++;
//$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $precio[1]);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['frec']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['hora']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['csemaca']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['cinicio']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['finicio']);$paz++;

$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['d1']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['d2']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['d3']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['d4']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['d5']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['d6']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['d7']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['d8']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['d9']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['d10']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['d11']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['d12']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['d13']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['d14']);$paz++;

$semanapasada= $r['d1']+$r['d2']+$r['d3']+$r['d4']+$r['d5']+$r['d6']+$r['d7'];
$seamaactual= $r['d8']+$r['d9']+$r['d10']+$r['d11']+$r['d12']+$r['d13']+$r['d14'];
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $semanapasada);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $seamaactual);$paz++;

$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['inscritos']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['nmetmat']);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['nmetmin']);$paz++;

$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['inicamp']);$paz++;
    if($r['ndiacamp']<0){
        $r['ndiacamp']=0;
    }
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['ndiacamp']);$paz++;
    if($r['dias_falta']<0){
        $r['dias_falta']=0;
    }
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['dias_falta']);$paz++;
$semanaant=round(($semanapasada/7),2);
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $semanaant);$paz++;
$ultdias=round(($seamaactual/7),2);
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial,$ultdias);$paz++;

$proy_faltante= round($r['dias_falta']*$ultdias,0);
$proy_fin_cam = $r['inscritos']+$proy_faltante;
$completa_meta = $r['nmetmat'] - $proy_fin_cam;
    if($completa_meta<0){
        $completa_meta=0;
    }
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $proy_faltante);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $proy_fin_cam);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $completa_meta);$paz++;
$objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$valorinicial, $r['observacion']);$paz++;
$color="FFFF4848";
if( $proy_fin_cam>=$r['nmetmat'] ){
    $color="FF35FF35";
}
elseif( $proy_fin_cam>=$r['nmetmin'] ){
    $color="FFFFFF48";
}
$objPHPExcel->getActiveSheet()->getStyle($az[($paz-3)].$valorinicial)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB($color);

}

    /*if($dcarrer!=''){
        $objPHPExcel->getActiveSheet()->mergeCells( "D".($valorinicial-$concarrera).":D".($valorinicial) );
        $objPHPExcel->getActiveSheet()->getStyle("D".($valorinicial-$concarrera).":D".($valorinicial))->applyFromArray($styleAlignmentLeft);
    }*/
$valorinicial++;
$objPHPExcel->getActiveSheet()->getRowDimension($valorinicial)->setRowHeight(25.25); // altura
    $objPHPExcel->getActiveSheet()->setCellValue("I".$valorinicial,"TOTALES");
    $objPHPExcel->getActiveSheet()->getStyle('I'.$valorinicial.":"."AK".$valorinicial)->applyFromArray($styleAlignmentRight);
    $objPHPExcel->getActiveSheet()->getStyle("A".$valorinicial.":"."AK".$valorinicial)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFD8E4BC');
    
    $col=9;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++; $col++; $col++; $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=SUM(".$az[$col].($valorinicial-$cont).":".$az[$col].($valorinicial-1).")"); $col++;
    array_push($sumatotal, $valorinicial);
    /*$objPHPExcel->getActiveSheet()->setCellValue( "B".($valorinicial-$cont), $fil );
    $objPHPExcel->getActiveSheet()->mergeCells( "B".($valorinicial-$cont).":B".($valorinicial-1) );
    $objPHPExcel->getActiveSheet()->setCellValue( "C".($valorinicial-$cont), $cins );
    $objPHPExcel->getActiveSheet()->mergeCells( "C".($valorinicial-$cont).":C".($valorinicial-1) );
    $objPHPExcel->getActiveSheet()->getStyle( "B".($valorinicial-$cont).":C".($valorinicial-1) )->applyFromArray($styleAlignment);
    */
$objPHPExcel->getActiveSheet()->getStyle('X6:Z'.$valorinicial)->applyFromArray($styleAlignmentRight);
$objPHPExcel->getActiveSheet()->getStyle("Z6".$valorinicial.":"."Z".$valorinicial)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFD8E4BC');
$objPHPExcel->getActiveSheet()->getStyle('A4:AK'.$valorinicial)->applyFromArray($styleThinBlackBorderAllborders);
//$objPHPExcel->getActiveSheet()->getStyle('AA4:AA'.$valorinicial)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);
////////////////////////////////////////////////////////////////////////////////////////////////
$valorinicial++;
$valorinicial++;
$objPHPExcel->getActiveSheet()->getRowDimension($valorinicial)->setRowHeight(25.25); // altura
    $objPHPExcel->getActiveSheet()->setCellValue("I".$valorinicial,"TOTALES");
    $objPHPExcel->getActiveSheet()->getStyle('I'.$valorinicial.":"."AK".$valorinicial)->applyFromArray($styleAlignmentRight);
    $objPHPExcel->getActiveSheet()->getStyle("I".$valorinicial.":"."AK".$valorinicial)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FFD8E4BC');
    $objPHPExcel->getActiveSheet()->getStyle("I".$valorinicial.":"."AK".$valorinicial)->applyFromArray($styleThinBlackBorderAllborders);
    
    $col=9;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++; $col++; $col++; $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;
        $objPHPExcel->getActiveSheet()->setCellValue($az[$col].$valorinicial,"=".$az[$col].implode($sumatotal,'+'.$az[$col])); $col++;


$objPHPExcel->getActiveSheet()->setTitle('Indice_Matricula');
// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);

/*// Redirect output to a client's web browser (Excel2007)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="Indice_MatriculaG_'.date("Y-m-d_H-i-s").'.xlsx"');
header('Cache-Control: max-age=0');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');*/
header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="Indice_MatriculaG_'.date("Y-m-d_H-i-s").'.xls"'); // file name of excel
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
