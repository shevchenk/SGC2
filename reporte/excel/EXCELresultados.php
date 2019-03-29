<?php
/*conexion*/
set_time_limit(3000);
ini_set('memory_limit','3072M');

//
//error_reporting(E_ALL);
//ini_set("display_errors", 1);


require_once '../../conexion/MySqlConexion.php';
require_once '../../conexion/configMySql.php';

/*crea obj conexion*/
$cn=MySqlConexion::getInstance();

$az=array(  'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z','AA','AB','AC','AD'
,'AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH'
,'BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ','CA','CB','CC','CD','CE','CF','CG','CH','CI','CJ','CK','CL'
,'CM','CN','CO','CP','CQ','CR','CS','CT','CU','CV','CW','CX','CY','CZ','DA','DB','DC','DD','DE','DF','DG','DH','DI','DJ','DK','DL','DM','DN','DO','DP'
,'DQ','DR','DS','DT','DU','DV','DW','DX','DY','DZ','EA','EB','EC','ED','EF','EG','EH','EI','EJ','EK','EL','EM','EN','EO','EP','EQ','ER','ES','ET','EU');
$azcount=array( 20,20,20,40,20,20,20,20,15,15,15,
    15,15,20,15,15,15,15,15,15,15,15,15,19,40,20,20,20,20,20,20,20,20,20,15,15
,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15
,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15
,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15
,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15
,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15
,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15);

$cingalu=$_GET['cingalu'];
$cgracpr=$_GET['cgracpr'];
$cusuari=$_GET['usuario'];
$alumno="";

$cfilial = str_replace(",","','",$_GET['cfilial']);
$cinstit =  str_replace(",","','",$_GET['cinstit']);
$ccarrer =  str_replace(",","','",$_GET['ccarrer']);

// si hay mas de 1 deberia mostrar todos los campos
$carreras = explode(",", $_GET['ccarrer']);
$limite = 7 * count($carreras);

$fechini = $_GET['fechini'];
$fechfin = $_GET['fechfin'];

$where='';
$order=" ORDER BY f.dfilial, ins.dinstit, c.dcarrer, p.dappape asc, p.dapmape asc , p.dnomper  asc ";

if ($cfilial) {
    $where .= " and g.cfilial in ('".$cfilial."') ";
}
if ($cinstit) {
    $where .= " and g.cinstit in ('".$cinstit."') ";
}

if ($ccarrer) {
    $where .= " and g.ccarrer in ('".$ccarrer."') ";
}


if($fechini!='' and $fechfin!=''){
    $where .=" AND date(g.finicio) between '".$fechini."' and '".$fechfin."' ";
}

$sql="
select
i.dcodlib inscripcion
,p.dappape
,p.dapmape
,p.dnomper
,'' asistencia
,pn.nota nota_final
,'' condicion
, f.dfilial filial
, ins.dinstit
, c.dcarrer carrer
, g.csemaca semestre
, g.finicio
from gracprp g
inner join conmatp co on co.cgruaca = g.cgracpr
inner join ingalum i on i.cingalu = co.cingalu
inner join personm p on p.cperson = i.cperson
inner join carrerm c on c.ccarrer = g.ccarrer
inner join modinga mo on mo.cmoding = i.cmoding
inner join filialm f on f.cfilial = g.cfilial
inner join instita ins on ins.cinstit = i.cinstit
left join posnota pn on pn.codlib = i.dcodlib
where 1 = 1
 ". $where . " " . $order;

$cn->setQuery($sql);
$control=$cn->loadObjectList();


// DATOS DE LA CARRERA PARA MOSTRARLO EN EL TITULO
$SQL = "select dcarrer from carrerm where ccarrer = '$ccarrer'";
$cn->setQuery($SQL);
$carrera = $cn->loadObjectList();


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

function colrow($az, $col, $row) {
    return $az[$col].$row;
}

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

// titulo principal
$az;
$row=3;
$col=0;
// EL TITULO SERA AGREGADO MANUALMENTE
$objPHPExcel->getActiveSheet()->setCellValue(colrow($az, $col, 1) , "RESULTADOS DEL EXAMEN DE ADMISION");
$objPHPExcel->getActiveSheet()->setCellValue(colrow($az, $col, 2) , "FACULTAD");
$objPHPExcel->getActiveSheet()->setCellValue(colrow($az, $col, 3) , "ESCUELA PROFESIONAL: ");
//$objPHPExcel->getActiveSheet()->getStyle(colrow($az, $col, $row))->getFont()->setSize(12);

//$objPHPExcel->getActiveSheet()->mergeCells('A2:M2');
//$objPHPExcel->getActiveSheet()->getStyle('A2:M2')->applyFromArray($styleAlignmentBold);

// fila titulo cabecera

$cabecera = array(
    "LIBRO DE CODIGO",
    "APELL PATERNO",
    "APELL MATERNO",
    "NOMBRES",
    "ASISTENCIA",
    "NOTA FINAL",
    "CONDICION",
    "FILIAL",
    "INSTITUCION",
    "CARRERA",
    "SEMETRE",
    "FECHA DE INICIO",
);
$row=4;
$col=0;

$colors = array(
    "FFDDD9C4",
    "FFC65911",
    "FFEBF1DE",
    "FF92D050",
    "FF8EA9DB",
    "FF3399FF",
);
$countColor = 0;
foreach($cabecera as $tit ) {
    if($col < $limite) {  // valida que no aparezcan otras columnas
        $objPHPExcel->getActiveSheet()->setCellValue(colrow($az, $col, $row), $tit);
        $objPHPExcel->getActiveSheet()->getStyle(colrow($az, $col, $row))->getAlignment()->setWrapText(true);
        $objPHPExcel->getActiveSheet()->getStyle(colrow($az, $col, $row))->applyFromArray($styleAlignmentBold);
        $objPHPExcel->getActiveSheet()->getColumnDimension($az[$col])->setWidth($azcount[$col]);
        $col++;
    }
}

//final columna de todo el excel
$finalCol = $col - 1;

// estilos para el titulo principal
//$objPHPExcel->getActiveSheet()->mergeCells(colrow($az, 0, 3) . ":" .  colrow($az, 8, 5));
//$objPHPExcel->getActiveSheet()->getStyle(colrow($az, 0, 3) . ":" .  colrow($az, 8, 5))->applyFromArray($styleAlignmentBold);

// rows body del excel

$row = 4;
$col = 0;
$cont = 0;
foreach($control As $r){
    $row++; // INICIA EN 5
    $paz=0; // columna
    $cont++;

    foreach ($r as $value)  {
        if ($paz < $limite){
            $objPHPExcel->getActiveSheet()->setCellValue($az[$paz].$row, $value); $paz++;
        }

    }
}

$objPHPExcel->getActiveSheet()->getStyle('A4:'.$az[$finalCol].$row)->applyFromArray($styleThinBlackBorderAllborders);
////////////////////////////////////////////////////////////////////////////////////////////////
$objPHPExcel->getActiveSheet()->setTitle('Resultados');
// Set active sheet index to the first sheet, so Excel opens this As the first sheet
$objPHPExcel->setActiveSheetIndex(0);

/*// Redirect output to a client's web browser (Excel2007)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="SGC.ReporteResultadoPostulantes.'.date("Y-m-d_His").'.xlsx"');
header('Cache-Control: max-age=0');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');*/
header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="ResultadoPostulante'.date("Y-m-d_H-i-s").'.xls"'); // file name of excel
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
