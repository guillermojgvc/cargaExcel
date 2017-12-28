package com.exco.cargaExcel.main;

import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;

import com.exco.cargaExcel.dbconnection.Conexion;
import com.exco.cargaExcel.documentos.excel.ValidarArchivoExcelLleno;
import com.exco.cargaExcel.model.Configuracion;
import com.exco.cargaExcel.model.InsertTemplate;
import com.exco.cargaExcel.model.PropiedadColumnaExcel;

public class App {

	public static void main(String[] args) {

		//System.out.println(Locale.getDefault());
		System.out.println("******** INICIA LA APLICACIÓN ********");
		long startTime = System.currentTimeMillis();
		List<PropiedadColumnaExcel> excelProperties = new ArrayList<PropiedadColumnaExcel>();
		InsertTemplate excelInsertTemplateProperties;
			
			
			Configuracion conf = new Configuracion();
			//excelProperties=conf.LeerArchivoPropiedades("/ExcelProperties.yml");
			//String configFileName = "ExcelProvinciaProperties.yml";
			//String configFileName = "ExcelCantonProperties.yml";
			String configFileName = "ExcelParroquiaProperties.yml";
			excelProperties=conf.LeerArchivoPropiedades("D:/Java/workspaces/workspaceTesis/cargaExcel/src/main/resources/" + configFileName);
			excelInsertTemplateProperties=conf.LeerArchivoPropiedadesInsertTemplate("D:/Java/workspaces/workspaceTesis/cargaExcel/src/main/resources/"+ configFileName);
			//excelProperties=conf.LeerArchivoPropiedades("/opt/appSITD/ExcelProvinciaProperties.yml");
			//excelInsertTemplateProperties=conf.LeerArchivoPropiedadesInsertTemplate("/opt/appSITD/ExcelProvinciaProperties.yml");
			if(excelProperties.isEmpty()){
				System.out.println("Lo sentimos el archivo de configuración no esta disponible");
				System.out.println("******** FIN DE LA APLICACIÓN ******** en : "
						+ (System.currentTimeMillis() - startTime) + "ms");
				return;
			}

			ValidarArchivoExcelLleno vae = new ValidarArchivoExcelLleno(excelProperties,excelInsertTemplateProperties);
			//String nombreArchivo = "D:/Guillermo/Senplades/HIT/PruebaCargaIndicador.xlsx";
			String fileName = "ICM_PARR";
			String nombreArchivoOrigen = "D:/Guillermo/Senplades/HIT/para carga/16 mayo_2016/"+fileName +".xlsx";
			String formattedDate = new SimpleDateFormat("dd-MM-yyyy_HH-mm-ss").format(new Date());
			String nombreArchivoDestino = "D:/Guillermo/Senplades/HIT/para carga/16 mayo_2016/"+fileName+"validado" + formattedDate + ".xlsx";
			
			/*Conexión EXCO
			 * Conexion conexion=new Conexion("oracle.jdbc.driver.OracleDriver", "jdbc:oracle:thin:@172.30.1.109:1521:orclexco", "DESCENTRALIZACION", "descentralizacion2015");*/
			
			/*Conexión Senplades Desarrollo*/
			Conexion conexion=new Conexion("oracle.jdbc.driver.OracleDriver", "jdbc:oracle:thin:@192.168.247.37:9859:INTERD", "DESCENTRALIZACION", "pR4cnACb");
			
			try {
				Connection con=conexion.conectar();
				List<Object> result;
				result = vae.validarArchivoXLSX(nombreArchivoOrigen, nombreArchivoDestino, true, con);
				if (Boolean.valueOf(result.get(0).toString())){
					System.out.println("Exito: " + result.get(1).toString());
				}else{
					System.out.println("Error: " + result.get(1).toString());
				}
				con.commit();
				con.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (ClassNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		
		System.out.println("******** FIN DE LA APLICACIÓN ******** en : "
				+ (System.currentTimeMillis() - startTime) + " ms");
		//System.out.println(Locale.getDefault());

	}
}
