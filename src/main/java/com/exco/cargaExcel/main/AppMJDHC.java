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

public class AppMJDHC {

	public static void main(String[] args) {

		//System.out.println(Locale.getDefault());
		System.out.println("******** INICIA LA APLICACIÓN ********");
		long startTime = System.currentTimeMillis();
		List<PropiedadColumnaExcel> excelProperties = new ArrayList<PropiedadColumnaExcel>();
		InsertTemplate excelInsertTemplateProperties;
			
			
			Configuracion conf = new Configuracion();
			//excelProperties=conf.LeerArchivoPropiedades("/ExcelProperties.yml");
			String configFileName = "ExcelPACL.yml";
			excelProperties=conf.LeerArchivoPropiedades("D:/Guillermo/Min Justicia/" + configFileName);
			excelInsertTemplateProperties=conf.LeerArchivoPropiedadesInsertTemplate("D:/Guillermo/Min Justicia/"+ configFileName);
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
			String fileName = "boletin_semanal_pacl_demo";
			String nombreArchivoOrigen = "D:/Guillermo/Min Justicia/BD/pruebas/"+fileName +".xlsx";
			String formattedDate = new SimpleDateFormat("dd-MM-yyyy_HH-mm-ss").format(new Date());
			String nombreArchivoDestino = "D:/Guillermo/Min Justicia/BD/pruebas/"+fileName+"validado" + formattedDate + ".xlsx";
			
			/*Conexión EXCO
			 * Conexion conexion=new Conexion("oracle.jdbc.driver.OracleDriver", "jdbc:oracle:thin:@172.30.1.109:1521:orclexco", "DESCENTRALIZACION", "descentralizacion2015");*/
			
			/*Conexión Senplades Desarrollo*/
			Conexion conexion=new Conexion("org.postgresql.Driver", "jdbc:postgresql://172.30.1.108:5432/MJDHC", "user", "user");
			
			try {
				Connection con=conexion.conectar();
				List<Object> result;
				result = vae.validarArchivoXLSX(nombreArchivoOrigen, nombreArchivoDestino, true, con);
				if (Boolean.valueOf(result.get(0).toString())){
					System.out.println("Exito: " + result.get(1).toString());
				}else{
					System.out.println("Error: " + result.get(1).toString());
				}
				//con.commit();
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
