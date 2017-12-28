package com.exco.cargaExcel.dbconnection;

import java.sql.*;

/**
 * Clase para la conexión de la BD
 * 
 * @author EXCO
 *
 */
public class Conexion {
	String driver ;
	String connectString; 
	String user;
	String password;
	
	/**
	 * Metodo para la obtener la conexión de la BD con el JBDC
	 * 
	 * @return Connection con la conexión a la BD
	 * @throws ClassNotFoundException
	 * @throws SQLException
	 */
	public Connection conectar() throws ClassNotFoundException, SQLException {
		Class.forName(driver);
		Connection con = DriverManager.getConnection(connectString, user,
				password);
		return con;
	}

	/**
	 * Constructor de la clase
	 *  
	 * @param driver String nombre del driver class ej: oracle.jdbc.driver.OracleDriver
	 * @param connectString String con la cadena de conexión ej: jdbc:oracle:thin:@172.30.1.109:1521:orcldemo 
	 * @param user String nombre de usuario
	 * @param password String password de usuario
	 */
	public Conexion(String driver, String connectString, String user,
			String password) {
		super();
		this.driver = driver;
		this.connectString = connectString;
		this.user = user;
		this.password = password;
	}
	
	
}