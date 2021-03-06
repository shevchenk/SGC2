<?
class servletGrupoAcademico extends controladorComandos{
	public function doPost(){
		$daoGrupoAcademico=creadorDAO::getGrupoAcademicoDAO();
    	switch ($_POST['accion']){
			case 'cargarGrupoAcademicoMatri':
				$data=array();
				$data['cmodali']=trim($_POST['cmodali']);
				$data['cfilial']=trim($_POST['cfilial']);
				$data['cinstit']=trim($_POST['cinstit']);
				$data['cciclo']=trim($_POST['cciclo']);
				$data['csemaca']=trim($_POST['csemaca']);
				if(isset($_POST['cinicio'])){
					$data['cinicio']=trim($_POST['cinicio']);
				}
                echo json_encode($daoGrupoAcademico->cargarGrupoAcademicoMatri($data));
			break;
            case 'cargar_grupo_academico':
				$data=array();
				$data['cmodali']=trim($_POST['cmodali']);
				$data['cfilial']=trim($_POST['cfilial']);
				$data['cinstit']=trim($_POST['cinstit']);
				$data['cciclo']=trim($_POST['cciclo']);
				$data['csemaca']=trim($_POST['csemaca']);
                echo json_encode($daoGrupoAcademico->cargarGrupoAcademico($data));
            break;
			case 'cargar_grupo_academico2':
				$data=array();
				$data['cfilial']=trim($_POST['cfilial']);
				$data['cinstit']=trim($_POST['cinstit']);
				$data['cciclo']=trim($_POST['cciclo']);
				$data['csemaca']=trim($_POST['csemaca']);
				$data['fechini']=trim($_POST['fechini']);
				$data['fechfin']=trim($_POST['fechfin']);
                echo json_encode($daoGrupoAcademico->cargarGrupoAcademico2($data));
            break;            
			case 'cargarAlumnos':
				$data=array();
				$data['id']=trim($_POST['id']);
                echo json_encode($daoGrupoAcademico->cargarAlumnos($data));
			break;
			case 'cargar_cursos_programados':
				$data=array();				
				$data['cgracpr']=trim($_POST['cgracpr']);
                echo json_encode($daoGrupoAcademico->cargarCursosProgramados($data));
            break;
			case 'guardar_grupos_academicos':
				$data=array();
				$data['ctipcar']=trim($_POST['ctipcar']);
				$data['cmodali']=trim($_POST['cmodali']);
				$data['ccarrer']=trim($_POST['ccarrer']);
				$data['ccurric']=trim($_POST['ccurric']);
				$data['cciclo']=trim($_POST['cciclo']);
				$data['cinicio']=trim($_POST['cinicio']);
				$data['csemaca']=trim($_POST['csemaca']);
				$data['cfilial']=trim($_POST['cfilial']);
				$data['cinstit']=trim($_POST['cinstit']);
				$data['cturno']=trim($_POST['cturno']);
				$data['chora']=trim($_POST['chora']);
				$data['observacion']=trim($_POST['observacion']);
				$data['dias']=trim($_POST['dias']);
				$data['usuario']=trim($_POST['usuario']);
				$data['cfilialx']=trim($_POST['cfilialx']);
				$data['finicio']=trim($_POST['finicio']);
				$data['ffinal']=trim($_POST['ffinal']);
				$data['nmetmat']=trim($_POST['nmetmat']);
				$data['nmetmin']=trim($_POST['nmetmin']);
				$data['fechafinsemestre']=trim($_POST['fechafinsemestre']);
				$data['fechainiciosemestre']=trim($_POST['fechainiciosemestre']);
				echo json_encode($daoGrupoAcademico->GuardarGruposAcademicos($data));
				break;
			case 'guardarPuntajePostulantes':
				$data=array();

				$data["data"]=trim($_POST['data']);
				$data['cusuari']=trim($_POST['cusuari']);
				$data['cfilial']=trim($_POST['cfilial']);

				echo json_encode($daoGrupoAcademico->guardarPuntajePostulantes($data));
				break;
			case 'ActualizarGrupoAcademico':
				$data=array();
				$data["cgruaca"]=trim($_POST['cgruaca']);				
				$data['cturno']=trim($_POST['cturno']);
				$data['chora']=trim($_POST['chora']);
				$data['dias']=trim($_POST['dias']);
				$data['usuario']=trim($_POST['usuario']);
				$data['cfilialx']=trim($_POST['cfilialx']);
				$data['finicio']=trim($_POST['finicio']);
				$data['ffinal']=trim($_POST['ffinal']);
				$data['nmetmat']=trim($_POST['nmetmat']);
				$data['nmetmin']=trim($_POST['nmetmin']);
				$data['observacion']=trim($_POST['observacion']);
				$data['fechafinsemestre']=trim($_POST['fechafinsemestre']);
				$data['fechainiciosemestre']=trim($_POST['fechainiciosemestre']);
				$data['valores']=trim($_POST['valores']);
				echo json_encode($daoGrupoAcademico->ActualizarGrupoAcademico($data));
      		break;
      		case 'cargarCursosAcademicos':
				$data=array();				
				$data['cgracpr']=trim($_POST['cgracpr']);				
                echo json_encode($daoGrupoAcademico->cargarCursosAcademicos($data));
            break;
            case 'cargarHorarioProgramado':
				$data=array();				
				$data['ccuprpr']=trim($_POST['ccuprpr']);
				$data['cdetgra']=trim($_POST['cdetgra']);
                echo json_encode($daoGrupoAcademico->cargarHorarioProgramado($data));
            break;
            case 'cargarDetalleGrupo':
				$data=array();
				$data['cgracpr']=trim($_POST['cgracpr']);				
                echo json_encode($daoGrupoAcademico->cargarDetalleGrupo($data));
            break;    
            case 'cargarDiasdelGrupo':
            	$data=array();
				$data['cgracpr']=trim($_POST['cgracpr']);				
                echo json_encode($daoGrupoAcademico->cargarDiasdelGrupo($data));
            break;
            case 'cargarCursosAcademicosAlumno':
            	$data=array();
				$data['cgracpr']=trim($_POST['cgracpr']);
				$data['cingalu']=trim($_POST['cingalu']);
                echo json_encode($daoGrupoAcademico->cargarCursosAcademicosAlumno($data));
            break;        
            case 'validarPasarRegistro':
            	$data=array();
				$data['ccuprpr']=trim($_POST['ccuprpr']);
				$data['gruequi']=trim($_POST['gruequi']);
				$data['cingalu']=trim($_POST['cingalu']);
                echo json_encode($daoGrupoAcademico->validarPasarRegistro($data));
            break;
            case 'crearPonderado':
            	$data=array();
				$data['cingalu']=trim($_POST['cingalu']);
                echo json_encode($daoGrupoAcademico->crearPonderado($data));
            break;
    		default:
                echo json_encode(array('rst'=>3,'msj'=>'Accion POST no encontrada'));
				break;
		}
	}
	public function doGet(){
		$daoGrupoAcademico=creadorDAO::getGrupoAcademicoDAO();
		switch ($_GET['action']){
			case 'cargarPostulantes':
				$data=array();
				$data['cfilial']=trim($_GET['cfilial']);
				$data['cinstit']=trim($_GET['cinstit']);
				$data['fechini']=trim($_GET['fechini']);
				$data['fechfin']=trim($_GET['fechfin']);
				echo json_encode($daoGrupoAcademico->cargarPostulantes($data));
				break;
			case 'jqgrid_grupo_academico':
					$page=$_GET["page"];
					$limit=$_GET["rows"];
					$sidx=$_GET["sidx"];
					$sord=$_GET["sord"];
		
					$where="";
					$param=array();
                                        
                    if( isset($_GET['dcarrer']) ) {
						if( trim($_GET['dcarrer'])!='' ) {
							$where.=" AND c.dcarrer like '%".trim($_GET['dcarrer'])."%' ";
						}
					}

					if( isset($_GET['dfilial']) ) {
						if( trim($_GET['dfilial'])!='' ) {
							$where.=" AND f.dfilial like '%".trim($_GET['dfilial'])."%' ";
						}
					}

					if( isset($_GET['dciclo']) ) {
						if( trim($_GET['dciclo'])!='' ) {
							$where.=" AND ci.dciclo like '%".trim($_GET['dciclo'])."%' ";
						}
					}

					if( isset($_GET['csemaca']) ) {
						if( trim($_GET['csemaca'])!='' ) {
							$where.=" AND g.csemaca like '%".trim($_GET['csemaca'])."%' ";
						}
					}

					if( isset($_GET['cinicio']) ) {
						if( trim($_GET['cinicio'])!='' ) {
							$where.=" AND g.cinicio like '%".trim($_GET['cinicio'])."%' ";
						}
					} 

					if( isset($_GET['finicio']) ) {
						if( trim($_GET['finicio'])!='' ) {
							$where.=" AND g.finicio='".trim($_GET['finicio'])."' ";
						}
					}                                        
										
					if(!$sidx)$sidx=1 ; 
		
					$row=$daoGrupoAcademico->JQGridCountGrupoAcademico( $where );
					$count=$row[0]['count'];
					if($count>0) {
							$total_pages=ceil($count/$limit);
					}else {
							$limit=0;
							$total_pages=0;
					}
		
					if($page>$total_pages) $page=$total_pages;
		
					$start=$page*$limit-$limit;
					
					$response=array("page"=>$page,"total"=>$total_pages,"records"=>$count);
					$data=$daoGrupoAcademico->JQGRIDRowsGrupoAcademico($sidx, $sord, $start, $limit, $where);
					$dataRow=array();
					$totalmatriculados=0;
					$vacantes=0;
					$indices=0;
					for($i=0;$i<count($data);$i++){
						array_push($dataRow, array("id"=>$data[$i]['id'],"cell"=>array( 								
								$data[$i]['dfilial'],
								$data[$i]['dcarrer'],
								$data[$i]['dciclo'],
								$data[$i]['csemaca'],
								$data[$i]['cinicio'],
                                $data[$i]['finicio'],
								$data[$i]['horario']								
								)
							)
						);
					}
					$response["rows"]=$dataRow;
					echo json_encode($response);
				break;
			default:
                echo json_encode(array('rst'=>3,'msj'=>'Accion GET no encontrada'));
				break;
		}	
	}
}
?>
