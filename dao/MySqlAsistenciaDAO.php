<?
class MySqlAsistenciaDAO{
    
    public function cargarAlumnos($cgrupo){
        $sql="
select 
i.cperson,
c.cingalu,
CONCAT_WS(' ',p.dappape,p.dapmape,p.dnomper) nombres,
s.seccion,
c.cgruaca
,CONCAT(p.ntelpe2,' | ',p.ntelper) telefono
,(select sum(nmonrec)
from recacap
where cingalu=c.cingalu
and cgruaca= c.cgruaca
and cconcep in (
select cconcep 
from concepp 
where (cctaing like '701.01%'
OR (cctaing like '701.03%' AND ccuota='1')
OR cctaing like '708.01%')
AND testfin='C'
)
GROUP BY cingalu)  conceptos

From conmatp c
inner join ingalum i on i.cingalu = c.cingalu
inner join personm p on p.cperson = i.cperson
left join seinggr s on s.cgrupo = c.cgruaca and s.cperson = i.cperson
where c.cgruaca ='$cgrupo' 
AND i.cestado='1' 
order by  conceptos desc ,p.dappape ASC;
            
        ";
        $db=creadorConexion::crear('MySql');
        $db->setQuery($sql);
        $data=$db->loadObjectList();
        if(count($data)>0){
            return array('rst'=>'1','msj'=>'Alumnos cargados','data'=>$data,'sql'=>$sql);
        }else{
            return array('rst'=>'2','msj'=>'No se encontraron alumnos','data'=>$data,'sql'=>$sql);
        }
    }

    public function rangoFechasGrupo($cgrupo){
        $db=creadorConexion::crear('MySql');
        $sql=" select CONCAT(finicio,'|',ffin) fgrupos ,DATE_FORMAT(now(),'%Y-%m-%d') fhoy ,if(finicio> now() , 0,1) estado ,cfrecue from gracprp g where g.cgracpr = '$cgrupo'; ";
        $db->setQuery($sql);
        $data=$db->loadObject();
        
        //OBTENIENDO LA FECHA DE LOS PRIMEROS 10 DIAS
        $fechas = $data->fgrupos;
        list($finicio,$ffin) = explode("|", $fechas);
        $data->finicio= $finicio;
        $fre = explode("-", $data->cfrecue);
        $dias = 0;
        $dfechas = array();
        while($dias < 15){
             
            $dd = date("w" , strtotime($finicio) );
            $dd++;
            if(in_array($dd, $fre)){
                $dias++;
                $dfechas[] = $finicio;
            }
            $fecha = date_create($finicio);
            date_add($fecha, date_interval_create_from_date_string('1 days'));
            $finicio = date_format($fecha, 'Y-m-d');
            
            
        }
       
        $data->ffin = $ffin;
        $data->fechas = $dfechas;
        
        $hoy = date("Y-m-d");
        $ayer = date("Y-m-d" , strtotime("-1 day"));
        $anteayer = date("Y-m-d" , strtotime("-2 day"));
        $debug = 1;
        $data->registrar = 0;
        foreach($dfechas as $i=>$v){
            if( $v == $hoy  ){
                $clase = $i+1;
                $data->registrar = 1;
                break;
            }elseif(strtotime($v) > strtotime($hoy) ){
                 $clase = $i;
                 $data->registrar = 0;
                 break;
            }
        }
        
        //$data->nroclase = $clase;
        $data->nroclase = 15 - $data->registrar;
        //$data->registrados = $clase - $data->registrar;
        $data->registrados = 15;
        $data->ayer = $ayer;
        $data->anteayer = $anteayer;
        
        
        if(count($data)>0){
            return array('rst'=>'1','msj'=>'fecha de rangos cargados','data'=>$data);
        }else{
            return array('rst'=>'2','msj'=>'No se encontraron fechas del grupo','data'=>"",'sql'=>$sql);
        }
    }//fin function rangoFechasGrupo

    function actualizarPosiciones() {
        $db=creadorConexion::crear('MySql');
        $set = 'SET SESSION group_concat_max_len = 1000000';
        $db->setQuery($set);
        $db->executeQuery();

        $sql = "select cgrupo, GROUP_CONCAT(id SEPARATOR '\',\'') ids
                from seinggr
                where cgrupo <> ''
                group by cgrupo";
        $db->setQuery($sql);
        $data=$db->loadObjectList();
        $db->iniciaTransaccion();
        foreach($data as $row) {
            $f= $this->rangoFechasGrupo($row["cgrupo"]);
            $fechas = $f["data"]->fechas;
            $cont = 0;
            foreach($fechas as $fecha) {
                $cont++;
                $sql = "update aluasist set posicion = $cont where fecha = '$fecha'  and idseing in ('".$row["ids"]."')";
                $db->setQuery($sql);

                if(!$db->executeQuery()){
                    $db->rollbackTransaccion();
                    return array('rst'=>'3','msj'=>'Error al Registrar Datos','sql'=>$sql);exit();
                }else if(!MySqlTransaccionDAO::insertarTransaccion($sql,"") ){
                    $db->rollbackTransaccion();
                    return array('rst'=>'3','msj'=>'Error al Registrar Datos','sql2'=>$sql);exit();
                }else{
                    $db->commitTransaccion();
                }
            }
        }
        return array('rst'=>'1','msj'=>'Datos registrados correctamente');exit();

    }

    
    //ACTUALIZA LA SECCION DE LOS alumnos de un grupo
    public function actualizarSeccionGrupo($post){
        $db=creadorConexion::crear('MySql');
        
        //PREGUNTA SI EXISTE
        $sql = "select * from seinggr as s where cingalu = '".$post["cingalu"]."' and cgrupo = '".$post["cgruaca"]."'";
        
        $db->setQuery($sql);
        $data=$db->loadObjectList();
        $deb = 1;
        //SI EXISTEN DATOS SE ACTUALIZA LA SECCION
        if(count($data)>0){
           $sql = "update seinggr set seccion='".$post["seccion"]."' where id = '".$data[0]["id"]."' ";
        }else{
            //SI NO EXISTE SE INSERTA LA SECCION
            $sql = "insert into seinggr set 
                    cingalu = '".$post["cingalu"]."' , 
                    cgrupo = '".$post["cgruaca"]."' ,
                    seccion='".$post["seccion"]."',
                    cperson = '".$post["cperson"]."',
                    cestado = '1',
                    fusuari = NOW(),
                    cusuari = '".$post["usuario"]."'
                ";
        }
        $db->iniciaTransaccion();
        $db->setQuery($sql);

        if(!$db->executeQuery()){
                $db->rollbackTransaccion();
                return array('rst'=>'3','msj'=>'Error al Registrar Datos','sql'=>$sql);exit();
        }else if(!MySqlTransaccionDAO::insertarTransaccion($sql,$r['cfilialx']) ){
                $db->rollbackTransaccion();
                return array('rst'=>'3','msj'=>'Error al Registrar Datos','sql2'=>$sql);exit();
        }else{
                $db->commitTransaccion();
                return array('rst'=>'1','msj'=>'Datos registrados correctamente');exit();
        }  
    }//fin function actualizarSeccionGrupo

    //LISTADO DE ALUMNOS DE UNA SECCION
    public function mostrarListadoCheck($cgrupo,$secc){

        $sql_posiciones = "";
        for ($i = 1; $i < 16; $i++) {
            $sql_posiciones .= " , (select a.estasist from aluasist a where  a.posicion = $i and a.idseing = s.id) pos$i ";
        }

        //OBTENIENDO ALUMNOS
        $sql = "select s.*,
                CONCAT_WS(' ',p.dappape,p.dapmape,p.dnomper) nombres
                ,CONCAT(p.ntelpe2,' | ',p.ntelper) telefono , i.dcodlib
                $sql_posiciones
                from seinggr s 
                inner join personm p on p.cperson = s.cperson
                inner join ingalum i on (i.cingalu=s.cingalu)
                where s.cgrupo = '$cgrupo' and s.seccion = '$secc' 
                AND i.cestado='1' 
				order by nombres";
        $db=creadorConexion::crear('MySql');
        $db->setQuery($sql);
        $data=$db->loadObjectList();
        if(count($data)>0){
            return array('rst'=>'1','msj'=>'Alumnos cargados','data'=>$data);
        }else{
            return array('rst'=>'2','msj'=>'No se encontraron alumnos en la seccion "'.$secc.'"','data'=>$data,'sql'=>$sql);
        }//fin consulta count(data)
    
    }//fin clase mostrarListadoCheck
    
    
    //REGISTRA LA ASISTENCIA DE UN ALUMNO
    public function registrarAsistencia($idse,$estado,$fecha,$post,$posicion){
        $db=creadorConexion::crear('MySql');
        
        //PREGUNTA SI EXISTE
        $sql = "select * from aluasist as a where a.idseing = '$idse' and fecha = '$fecha'";
        
        $db->setQuery($sql);
        $data=$db->loadObject();
        $deb = 1;
        //SI EXISTEN DATOS SE ACTUALIZA LA SECCION
        if(count($data)>0){
           $sql = "update aluasist set 
               estasist='$estado' ,
                   fusuari = NOW(),
                    cusuari = '".$post["usuario"]."'
               where idaluasi = '".$data->idaluasi."' ";
        }else{
            //SI NO EXISTE SE INSERTA LA SECCION
            $sql = "insert into aluasist set 
                    idseing = '".$idse."' , 
                    estasist = '".$estado."' ,
                    fecha='".$fecha."',
                    posicion = $posicion,
                    fusuari = NOW(),
                    cusuari = '".$post["usuario"]."'
                ";
        }
        $db->iniciaTransaccion();
        $db->setQuery($sql);

        if(!$db->executeQuery()){
                $db->rollbackTransaccion();
                return array('rst'=>'3','msj'=>'Error al Registrar Datos','sql'=>$sql);exit();
        }else if(!MySqlTransaccionDAO::insertarTransaccion($sql,$r['cfilialx']) ){
                $db->rollbackTransaccion();
                return array('rst'=>'3','msj'=>'Error al Registrar Datos','sql2'=>$sql);exit();
        }else{
                $db->commitTransaccion();
                return array('rst'=>'1','msj'=>'Datos registrados correctamente');exit();
        }  
    }//fin function registrarAsistencia
 
    
    //LISTADO DE LA ASISTENCIA DE UN ALUMNO DE UNA FECHA DETERMINADA
    public function asistenciaAlumno($id ,$f){
        
        //OBTENIENDO ALUMNOS
        $sql = "select estasist as flat from aluasist where idseing = '$id' and fecha = '$f'";
        $db=creadorConexion::crear('MySql');
        $db->setQuery($sql);
        $data=$db->loadObject();
        if(count($data)>0){
            return array('rst'=>'1','msj'=>'Asistencia cargada','data'=>$data->flat);
        }else{
            return array('rst'=>'1','msj'=>'Asistencia cargada','data'=>0);
            //return array('rst'=>'2','msj'=>'No se encontraron alumnos en la seccion "'.$secc.'"','data'=>$data,'sql'=>$sql);
        }//fin consulta count(data)
    
    }//fin clase asistenciaAlumno
    
}
?>
