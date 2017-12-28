package com.exco.cargaExcel.documentos.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Types;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.exco.cargaExcel.enums.Obligatoriedad;
import com.exco.cargaExcel.enums.TipoDato;
import com.exco.cargaExcel.model.InsertTemplate;
import com.exco.cargaExcel.model.PropiedadColumnaExcel;

/**
 * Clase para la validación de un archivo de Excel en Java usando Apache POI
 * 
 * @author Guillermo Vaca - EXCO
 * @Mail guillermov@ecorporativa.com
 *
 */
public class ValidarArchivoExcelLlenoRespaldo {

	// Listado proveniente de la lectura del archivo "Excel<*>Properties.yml"
	List<PropiedadColumnaExcel> excelProperties;
	// Template de la sentencia INSERT en la BD
	InsertTemplate excelInsertTemplateProperties;
	// Mapa de la columnas y los valores unicos (distinct) de cada columna para
	// las consultas en la BD
	Map<Integer, LinkedHashSet<String>> mapeoFks = new LinkedHashMap<Integer, LinkedHashSet<String>>();
	// Mapa de la columnas y los valores unicos (distinct) de cada columna para
	// las consultas en la BD y verificar la existencia del ID
	Map<Integer, LinkedHashSet<String>> mapeoIds = new LinkedHashMap<Integer, LinkedHashSet<String>>();
	// Variable de texto para la concatenación de los mensajes de validación del
	// archivo
	String observacion;

	/**
	 * Constructor vacío de la clase
	 */
	public ValidarArchivoExcelLlenoRespaldo() {
		super();
	}

	/**
	 * Constructor con el listado de propiedades obtenidas del archivo
	 * "Excel<*>Properties.yml"
	 * 
	 * @param excelProperties
	 */
	public ValidarArchivoExcelLlenoRespaldo(List<PropiedadColumnaExcel> excelProperties) {
		super();
		this.excelProperties = excelProperties;
	}

	/**
	 * Constructor con el listado de propiedades y el InsertTemplate obtenidos
	 * del archivo "Excel<*>Properties.yml"
	 * 
	 * @param excelProperties
	 *            Lista de objetos PropiedadColumnaExcel obtenido del yaml tag
	 *            "Columnas"
	 * @param excelInsertTemplateProperties
	 *            Lista de objetos InsertTemplate obtenido del yaml tag
	 *            "InsertTemplate"
	 */
	public ValidarArchivoExcelLlenoRespaldo(
			List<PropiedadColumnaExcel> excelProperties,
			InsertTemplate excelInsertTemplateProperties) {
		super();
		this.excelProperties = excelProperties;
		this.excelInsertTemplateProperties = excelInsertTemplateProperties;
	}

	/**
	 * Metodo para validación de un archivo de excel provisto para la carga que
	 * retorna un valor booleano de si el archivo es válido o no
	 * 
	 * @param nombreArchivo
	 *            String con la ubicación física de un archivo excel xlsx a
	 *            cargar
	 * @param primeraFilaCabecera
	 *            Booleano que identifica si el archivo de excel debe ser
	 *            procesado desde la fila (1) true o (0) false
	 * @param connection
	 *            Connection instancia a una conexión JDBC para la consulta y
	 *            carga de los valores de excel
	 * @return Lista de Objetos donde el indice 0 es un booleano true si un
	 *         archivo es válido con la definición del archivo de configuración
	 *         o false en caso contrario y el indice 1 es el mensaje del
	 *         procesamiento del archivo
	 * @throws IOException
	 * @throws SQLException
	 */
	@SuppressWarnings("rawtypes")
	public List<Object> validarArchivoXLSX(String nombreArchivoOrigen,
			String nombreArchivoDestino, boolean primeraFilaCabecera,
			Connection connection) throws IOException, SQLException {
		List<Object> resultObj = new ArrayList<Object>();
		// Mensaje para mostrar al usuario en una página JSF
		String mensaje = "";
		// Variable de validación del archivo Excel
		boolean valido = true;
		// Texto de la cabecera de la columna de excel
		String encabezado;
		// Variable de la filas a leer
		XSSFRow row;
		// Variable de la celda a leer
		XSSFCell cell;
		// Arreglo de observaciones para el archivo de Excel
		String[] vectorObservaciones = null;
		// Lista de listas de objetos con los valores de los inserts de cada
		// fila
		List<List<Object>> vectorInserts = null;
		// Colección de hash con el listado de campos a buscar el ID
		LinkedHashSet<String> listadoUnicoId = null;
		// Mapeo de hash con los valor <fila,"campo">
		LinkedHashMap<Integer, String> mapaUnicoId = null;
		// Mapeo de hash con los valor <columna,<fila,"campo">>
		LinkedHashMap<Integer, LinkedHashMap<Integer, String>> listadoMapaUnicoFk = null;
		// Mapeo de hash con los valor <columna,<fila,"campo">>
		LinkedHashMap<Integer, LinkedHashMap<Integer, String>> listadoMapaUnicoId = null;
		// Variable que asigna la propiedad de Excel de acuerdo al indice de la
		// lista excelProperties.
		PropiedadColumnaExcel propiedadColumnaExcel;
		// Variable para identificar si es una columna con valor a buscar el Fk
		boolean isFk;
		// Variable para identificar si es una columna con valor a buscar el Id
		boolean isId;

		// Lectura del archivo Excel
		InputStream ExcelFileToRead = new FileInputStream(nombreArchivoOrigen);

		// Apertura del archivo con formato XLSX y asignación de la localización
		// (es_EC = Ecuador)
		XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
		// Locale.setDefault(new Locale("es_EC"));

		// Obtener la primera hoja del archivo Excel a leer
		XSSFSheet sheet = wb.getSheetAt(0);

		// Lectura del número de filas
		int numeroFilas = sheet.getLastRowNum();
		// Lectura del número de columnas
		int numeroColumnas = sheet.getRow(0).getLastCellNum();

		mensaje+=("Filas encontradas: " + numeroFilas + "\n");
		mensaje+=("Columnas encontradas: " + numeroColumnas + "\n");
		System.out.println("******** Filas encontradas: " + numeroFilas);
		System.out.println("******** Columnas encontradas: " + numeroColumnas);

		// Seteo del tamaño al vector y la cabecera
		vectorObservaciones = new String[numeroFilas + 1];
		if (primeraFilaCabecera) {
			vectorObservaciones[0] = "Observaciones";
			// Iterador de celdas, para eliminar los estilos de la celda
			// cabecera
			Iterator cells = sheet.getRow(0).cellIterator();

			while (cells.hasNext()) {
				cell = (XSSFCell) cells.next();
				if (cell != null) {
					cell.setCellStyle(null);
				}
				cell.setCellType(XSSFCell.CELL_TYPE_STRING);
			}
		}

		// Seteo del tamaño de la lista de inserts
		vectorInserts = new ArrayList<List<Object>>();
		for (int i = 0; i < numeroFilas + 1; i++) {
			vectorInserts.add(new ArrayList<Object>());
		}

		// Verificación de las columnas Excel concuerden con el listado de
		// propiedades
		if (excelProperties.size() == numeroColumnas) {
			listadoMapaUnicoFk = new LinkedHashMap<Integer, LinkedHashMap<Integer, String>>();
			listadoMapaUnicoId = new LinkedHashMap<Integer, LinkedHashMap<Integer, String>>();
			for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
				isFk = false;
				isId = false;
				// Lectura de la propiedad de Excel en el orden del archivo de
				// configuraciones
				propiedadColumnaExcel = excelProperties.get(i);

				// Verificación si la propiedad de Excel identifica a la columna
				// como FK y asignación de variables asociadas
				if (propiedadColumnaExcel.getTipo().equals(
						TipoDato.Fk.toString())) {
					listadoUnicoId = new LinkedHashSet<String>();
					mapaUnicoId = new LinkedHashMap<Integer, String>();
					isFk = true;
				}

				// Verificación si la propiedad de Excel identifica a la columna
				// como ID
				if (propiedadColumnaExcel.getTipo().equals(
						TipoDato.Id.toString())) {
					listadoUnicoId = new LinkedHashSet<String>();
					mapaUnicoId = new LinkedHashMap<Integer, String>();
					isId = true;
				}

				// Asigna el encabezado según sea el caso
				if (primeraFilaCabecera) {
					encabezado = sheet.getRow(0).getCell(i)
							.getStringCellValue();
				} else {
					encabezado = null;
				}

				// Iteración de las filas "j" pertenecientes a la columna "i"
				for (int j = primeraFilaCabecera == true ? 1 : 0; j < numeroFilas + 1; j++) {
					row = sheet.getRow(j);
					cell = row.getCell(i);

					// Elimina los estilos a las celdas no nulas
					if (cell != null) {
						cell.setCellStyle(null);
					}

					// Evalua si una celda es válida y asigna una observación en
					// caso de no serlo
					if (!validarCelda(
							cell,
							propiedadColumnaExcel.getObligatoriedad().equals(
									Obligatoriedad.Requerido.mostrar()) ? Obligatoriedad.Requerido
									: Obligatoriedad.Opcional,
							propiedadColumnaExcel, encabezado, isFk, isId,
							listadoUnicoId, mapaUnicoId, j,
							vectorInserts.get(j))) {
						if (!observacion.isEmpty()) {
							valido = false;
							if (vectorObservaciones[row.getRowNum()] != null) {
								// Concatena la observación actual con la
								// observación inicial
								vectorObservaciones[row.getRowNum()] = vectorObservaciones[row
										.getRowNum()]
										.concat(", " + observacion);
							} else {
								// Asigna la observación inicial
								vectorObservaciones[row.getRowNum()] = (observacion);
							}
						}
					}

				}

				// Asigna el listado de ID únicos a consultar a la columna "i" y
				// el mapeo <columna,<fila,"campo">>
				if (isFk) {
					mapeoFks.put(i, listadoUnicoId);
					listadoMapaUnicoFk.put(i, mapaUnicoId);
				}

				if (isId) {
					mapeoIds.put(i, listadoUnicoId);
					listadoMapaUnicoId.put(i, mapaUnicoId);
				}

			}

			// Busqueda del ID foráneo si el archivo es válido, caso contrario
			// imprime las observaciones en un nuevo archivo de Excel
			if (valido) {
				System.out
						.println("******** Inicia la busqueda de los ID foraneos");
				LinkedHashMap<Integer, String> aux = null;
				LinkedHashMap<String, String> resultado = null;

				// Iteración del mapeo de ID's <columna,<id1,id2,...,idn>>
				for (Map.Entry<Integer, LinkedHashSet<String>> entry : mapeoIds
						.entrySet()) {

					aux = new LinkedHashMap<Integer, String>(
							listadoMapaUnicoId.get(entry.getKey()));
					resultado = new LinkedHashMap<String, String>();

					Iterator it = entry.getValue().iterator();
					String query = new String(excelProperties.get(
							entry.getKey()).getNativeQuery());
					// Iteracion de los valores mapeados
					while (it.hasNext()) {
						String value = (String) it.next();
						// Consulta ID
						resultado.put(value,
								consultaExistenciaID(connection, query, value));
					}

					for (int i = primeraFilaCabecera == true ? 1 : 0; i <= aux
							.size(); i++) {
						row = sheet.getRow(i);
						cell = row.getCell(entry.getKey());

						/**
						 * // Valor obtenido de la consulta a la BD, descomentar
						 * la // siguiente línea si se desea reemplazar los
						 * valores // actuales con los nuevos
						 * cell.setCellValue(resultado.get(aux.get(i)));
						 **/

						if (resultado.get(aux.get(i)) == null) {
							valido = false;
							concatenarObservacion(vectorObservaciones,
									(primeraFilaCabecera ? sheet.getRow(0)
											.getCell(entry.getKey())
											.getStringCellValue() : null),
									"no existe ID en la Base de datos", i);
						}

					}

				}

				// Iteración del mapeo de Fk's <columna,<id1,id2,...,idn>>
				for (Map.Entry<Integer, LinkedHashSet<String>> entry : mapeoFks
						.entrySet()) {

					aux = new LinkedHashMap<Integer, String>(
							listadoMapaUnicoFk.get(entry.getKey()));
					resultado = new LinkedHashMap<String, String>();

					Iterator it = entry.getValue().iterator();
					String query = new String(excelProperties.get(
							entry.getKey()).getNativeQuery());
					// Iteracion de los valores mapeados
					while (it.hasNext()) {
						String value = (String) it.next();
						// Consulta ID
						resultado.put(value,
								consultaID(connection, query, value));
					}

					for (int i = primeraFilaCabecera == true ? 1 : 0; i <= aux
							.size(); i++) {
						row = sheet.getRow(i);
						cell = row.getCell(entry.getKey());

						/**
						 * // Valor obtenido de la consulta a la BD, descomentar
						 * la // siguiente línea si se desea reemplazar los
						 * valores // actuales con los nuevos
						 * cell.setCellValue(resultado.get(aux.get(i)));
						 **/

						if (resultado.get(aux.get(i)) == null) {
							valido = false;
							concatenarObservacion(
									vectorObservaciones,
									(primeraFilaCabecera ? sheet.getRow(0)
											.getCell(entry.getKey())
											.getStringCellValue() : null),
									"no existe Fk relacionado al valor buscado",
									i);
						} else {
							// Asignar el valor de la consulta a los valores de
							vectorInserts.get(i)
									.set(entry.getKey(),
											Integer.parseInt(resultado.get(aux
													.get(i))));
						}

					}

				}
			}

			// Si el archivo no es válido empieza la generación del archivo de
			// validado
			if (!valido) {
				System.out
						.println("******** Inicia la generación del archivo validado");
				mensaje+=("Se ha generado el archivo: \n"
						+ nombreArchivoDestino
						+ "\n con las observaciones de validación del mismo");
				for (int i = 0; i < vectorObservaciones.length; i++) {
					if (vectorObservaciones[i] != null) {
						sheet.getRow(i).createCell(numeroColumnas)
								.setCellValue(vectorObservaciones[i]);
					}
				}
			} else {
				System.out.println("******** Inicia la carga en batch");
				valido = insertBatch(connection, primeraFilaCabecera,
						vectorInserts);
				mensaje+=("Se ha insertado el archivo Excel en la base de datos");
			}
		} else {
			System.out
					.println("El número de columnas no concuerda con el archivo de configuración");
			mensaje+=("El número de columnas no concuerda con el archivo de configuración");
		}

		FileOutputStream fileOut = new FileOutputStream(nombreArchivoDestino);
		// FileOutputStream fileOut = new
		// FileOutputStream("/opt/appSITD/Validado" + formattedDate + ".xlsx");

		// write this workbook to an Outputstream.
		wb.write(fileOut);
		wb.close();
		resultObj.add(valido);
		resultObj.add(mensaje);
		return resultObj;
	}

	/**
	 * Método para concatenar las observaciones que serán asignadas al archivo
	 * de validación
	 * 
	 * @param vectorObservaciones
	 * @param encabezado
	 * @param mensaje
	 * @param fila
	 */
	public void concatenarObservacion(String[] vectorObservaciones,
			String encabezado, String mensaje, int fila) {

		observacion = "Campo " + encabezado + " " + mensaje;
		if (vectorObservaciones[fila] != null) {
			// Concatena la observación actual con la observación inicial
			vectorObservaciones[fila] = vectorObservaciones[fila].concat(", "
					+ observacion);
		} else {
			// Asigna la observación inicial
			vectorObservaciones[fila] = (observacion);
		}

	}

	/**
	 * Método para la inserción en batch (bulk insert) de cada fila de un
	 * archivo excel
	 * 
	 * @param connection
	 *            Connection instancia a una conexión JDBC para la consulta y
	 *            carga de los valores de excel
	 * @param primeraFilaCabecera
	 *            Booleano que identifica si el archivo de excel debe ser
	 *            procesado desde la fila (1) true o (0) false
	 * @param vectorInserts
	 *            Lista de lista de objetos correspodientes a cada valor de cada
	 *            fila del archivo excel agregados por el tipo de dato definido
	 *            en el archivo de configuración
	 * @return Booleano true si los inserts se completaron de forma exitosa,
	 *         caso contrario retorna un false
	 */
	public boolean insertBatch(Connection connection,
			boolean primeraFilaCabecera, List<List<Object>> vectorInserts) {
		String query = this.excelInsertTemplateProperties.getInsert();
		PreparedStatement ps;
		try {
			ps = connection.prepareStatement(query);

			final int batchSize = 1000;

			for (int i = primeraFilaCabecera == true ? 1 : 0; i < vectorInserts
					.size(); i++) {
				for (int j = 0; j < vectorInserts.get(i).size(); j++) {
					switch (TipoDato.valueOf(excelProperties.get(j).getTipo())) {
					case Entero:
						if (vectorInserts.get(i).get(j) != null) {
							ps.setInt(j + 1, (Integer) vectorInserts.get(i)
									.get(j));
						} else {
							ps.setNull(j + 1, Types.INTEGER);
						}
						break;

					case Doble:
						if (vectorInserts.get(i).get(j) != null) {
							ps.setDouble(j + 1, (Double) vectorInserts.get(i)
									.get(j));
						} else {
							ps.setNull(j + 1, Types.DOUBLE);
						}
						break;

					case Texto:
						if (vectorInserts.get(i).get(j) != null) {
							ps.setString(j + 1, vectorInserts.get(i).get(j)
									.toString());
						} else {
							ps.setString(j + 1, null);
						}
						break;

					case Fk:
						if (vectorInserts.get(i).get(j) != null) {
							ps.setInt(j + 1, (Integer) vectorInserts.get(i)
									.get(j));
						} else {
							ps.setNull(j + 1, Types.INTEGER);
						}
						break;

					case Id:
						if (vectorInserts.get(i).get(j) != null) {
							ps.setInt(j + 1, (Integer) vectorInserts.get(i)
									.get(j));
						} else {
							ps.setNull(j + 1, Types.INTEGER);
						}
						break;

					default:
						return false;
					}
				}

				ps.addBatch();
				if (i % batchSize == 0) {
					ps.executeBatch();
				}
			}

			ps.executeBatch(); // insert remaining records
			ps.close();

		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return false;
		}

		return true;
	}

	/**
	 * Método para la consulta de las claves foraneas pertenecientes a los
	 * campos marcados como ID en el archivo de configuración previo a un
	 * filtrado de unicidad
	 * 
	 * @param connection
	 *            Connection instancia a una conexión JDBC para la consulta y
	 *            carga de los valores de excel
	 * @param query
	 *            String con el query por defecto para la busqueda del ID
	 *            definido en el archivo de configuración
	 * @param value
	 *            String con el valor a consultar (parametro para el
	 *            preparedstatement)
	 * @return String retorna el primer elemento de la consulta o null en caso
	 *         de no existir un ID.
	 * @throws SQLException
	 */
	private String consultaID(Connection connection, String query, String value)
			throws SQLException {
		String resultado = null;
		PreparedStatement preparedStatement = null;
		preparedStatement = connection.prepareStatement(query);
		preparedStatement.setString(1, value);
		try {
			ResultSet rs = preparedStatement.executeQuery();
			try {
				while (rs.next()) {
					resultado = rs.getString(1);
				}
			} finally {
				try {
					rs.close();
				} catch (Exception ignore) {
				}
			}
		} finally {
			try {
				preparedStatement.close();
			} catch (Exception ignore) {
			}
		}
		return resultado;
	}

	/**
	 * Método para la consulta de las claves primarias pertenecientes a los
	 * campos marcados como ID en el archivo de configuración
	 * 
	 * @param connection
	 *            Connection instancia a una conexión JDBC para la consulta y
	 *            carga de los valores de excel
	 * @param query
	 *            String con el query por defecto para la busqueda del ID
	 *            definido en el archivo de configuración
	 * @param value
	 *            String con el valor a consultar (parametro para el
	 *            preparedstatement)
	 * @return String retorna el primer elemento de la consulta o null en caso
	 *         de no existir un ID.
	 * @throws SQLException
	 */
	private String consultaExistenciaID(Connection connection, String query,
			String value) throws SQLException {
		String resultado = null;
		PreparedStatement preparedStatement = null;
		preparedStatement = connection.prepareStatement(query);
		// preparedStatement.setString(1, value);
		preparedStatement.setInt(1, Integer.valueOf(value));
		try {
			ResultSet rs = preparedStatement.executeQuery();
			try {
				while (rs.next()) {
					resultado = rs.getString(1);
				}
			} finally {
				try {
					rs.close();
				} catch (Exception ignore) {
				}
			}
		} finally {
			try {
				preparedStatement.close();
			} catch (Exception ignore) {
			}
		}
		return resultado;
	}

	/**
	 * Método para la evaluación de expresiones regulares proveniente del
	 * archivo de configuración en las celdas de tipo TEXTO.
	 * 
	 * @param regex
	 *            String con la expresión regular a evaluar
	 * @param cell
	 *            XSSFCell celda de Excel con las propiedades y atributos de
	 *            dicha celda
	 * @param encabezado
	 *            String con el nombre de la cabecera de la columna, este
	 *            encabezado sirve para asignar el campo observación
	 * @return Boolean true si la celda de Excel tipo TEXTO cumple con la
	 *         expresión regular, caso contrario false
	 */
	private boolean evaluarExpresionTexto(String regex, XSSFCell cell,
			String encabezado) {
		boolean expresionEvaluada = cell.getStringCellValue().matches(regex);
		if (expresionEvaluada == false)
			observacion = "Campo " + encabezado
					+ " no concuerda con el formato";
		return expresionEvaluada;
	}

	/**
	 * Método para la evaluación de expresiones regulares proveniente del
	 * archivo de configuración en las celdas de tipo ENTERO.
	 * 
	 * @param regex
	 *            String con la expresión regular a evaluar
	 * @param cell
	 *            XSSFCell celda de Excel con las propiedades y atributos de
	 *            dicha celda
	 * @param encabezado
	 *            String con el nombre de la cabecera de la columna, este
	 *            encabezado sirve para asignar el campo observación
	 * @return Boolean true si la celda de Excel tipo ENTERO cumple con la
	 *         expresión regular, caso contrario false
	 */
	private boolean evaluarExpresionNumeroEntero(String regex, XSSFCell cell,
			String encabezado) {
		boolean expresionEvaluada = cell.getRawValue().matches(regex);
		if (expresionEvaluada == false)
			observacion = "Campo " + encabezado
					+ " no concuerda con el formato";
		return expresionEvaluada;
	}

	/**
	 * Método para la evaluación de expresiones regulares proveniente del
	 * archivo de configuración en las celdas de tipo ENTERO.
	 * 
	 * @param regex
	 *            String con la expresión regular a evaluar
	 * @param cell
	 *            XSSFCell celda de Excel con las propiedades y atributos de
	 *            dicha celda
	 * @param decimalPlaces
	 *            Integer con el valor para redondeo de la presición de
	 *            decimales calculada en la celda de Excel
	 * @param encabezado
	 *            String con el nombre de la cabecera de la columna, este
	 *            encabezado sirve para asignar el campo observación
	 * @return Boolean true si la celda de Excel tipo ENTERO cumple con la
	 *         expresión regular, caso contrario false
	 */
	private boolean evaluarExpresionNumeroDouble(String regex, XSSFCell cell,
			int decimalPlaces, String encabezado) {
		DecimalFormat df = new DecimalFormat("#");
		int fractionalDigits = decimalPlaces; // say 2
		df.setMaximumFractionDigits(fractionalDigits);
		boolean expresionEvaluada = df.format(cell.getNumericCellValue())
				.matches(regex);
		if (expresionEvaluada == false)
			observacion = "Campo " + encabezado
					+ " no concuerda con el formato";
		return expresionEvaluada;
	}

	/**
	 * Método para agregar los valores de inserción en la lista de inserción
	 * 
	 * @param tipoDato
	 *            String con el tipo valor del tipo de dato proveniente del
	 *            archivo de confiuguración que será comparado con el enumerador
	 *            TipoDato
	 * @param cell
	 *            XSSFCell celda Excel con las propiedades y atributos de dicha
	 *            celda
	 * @param decimalPlaces
	 *            Integer con el valor para redondeo de la presición de
	 *            decimales calculada en la celda de Excel para permitir la
	 *            inserción de datos según la presición de la BD
	 * @param insertValues
	 *            Lista de objetos que agrega valores de los campos a insertar
	 *            en la BD según el InsertTemplate
	 */
	public void insertAddValues(String tipoDato, XSSFCell cell,
			int decimalPlaces, List<Object> insertValues) {
		DecimalFormat df = new DecimalFormat("#");
		if (cell != null) {
			switch (TipoDato.valueOf(tipoDato)) {
			case Entero:
				insertValues.add((int) cell.getNumericCellValue());
				break;

			case Id:
				insertValues.add((int) cell.getNumericCellValue());
				break;

			case Doble:
				if (cell.getRawValue() == null) {
					insertValues.add(null);
				} else {
					int fractionalDigits = decimalPlaces;
					df.setMaximumFractionDigits(fractionalDigits);
					// System.out.println(Double.parseDouble(df.format(cell.getNumericCellValue())));
					/*
					 * insertValues.add(Double.parseDouble(df.format(cell
					 * .getNumericCellValue())));
					 */
					insertValues.add(cell.getNumericCellValue());
				}
				break;

			case Texto:
				insertValues.add(cell.getStringCellValue());
				break;

			case Fk:
				insertValues.add(null);
				break;

			default:
				System.out.println("No valido con el archivo de configuración");
				break;
			}
		} else {
			insertValues.add(null);
		}
	}

	/**
	 * Método para validar la celda de Excel, si es obligatorio o si cumple con
	 * la expresión regular requerida en el archivo de configuración
	 * 
	 * @param cell
	 *            XSSFCell celda Excel con las propiedades y atributos de dicha
	 *            celda
	 * @param obligatoriedad
	 *            Enum Obligatoriedad, permite comparar si una celda es
	 *            Obligatoria (R) u Opcional (O). Estos valores se obtienen del
	 *            archivo de configuración
	 * @param propiedadColumnaExcel
	 *            PropiedadColumnaExcel son las propiedades obtenidas para esa
	 *            celda del archivo de configuración, contiene expresiones
	 *            regulares y parámetros de validación
	 * @param encabezado
	 *            String con el nombre de la cabecera de la columna, este
	 *            encabezado sirve para asignar el campo observación
	 * @param isId
	 *            Booleano que define si una celda es del tipo ID y se requiere
	 *            consultar su ID
	 * @param listadoUnicoId
	 *            Lista de hash con el listado único de los ID a consultar a la
	 *            BD
	 * @param mapaUnicoId
	 *            Mapeo del número de fila y el String único de la columna para
	 *            la consulta en la BD y su fácil localización
	 * @param pos
	 *            Integer con la posición de la fila para el mapeo.
	 * @param insertValues
	 *            Lista de objetos que agrega valores de los campos a insertar
	 *            en la BD según el InsertTemplate
	 * @return Booleano true en caso de que la celda no presente observaciones a
	 *         la validación, false en caso contrario
	 */
	private boolean validarCelda(XSSFCell cell, Obligatoriedad obligatoriedad,
			PropiedadColumnaExcel propiedadColumnaExcel, String encabezado,
			boolean isFk, boolean isId, LinkedHashSet<String> listadoUnicoId,
			LinkedHashMap<Integer, String> mapaUnicoId, int pos,
			List<Object> insertValues) {
		boolean estado = true;
		observacion = new String();

		if (cell == null) {
			if (obligatoriedad.equals(Obligatoriedad.Requerido)) {
				estado = false;
				observacion = "Campo " + encabezado + " no debe ser nulo";
			}
			insertAddValues(propiedadColumnaExcel.getTipo(), cell, 0,
					insertValues);
			return estado;
		}
		switch (cell.getCellType()) {
		case XSSFCell.CELL_TYPE_STRING:
			cell.setCellValue(cell.getStringCellValue().trim());
			if (obligatoriedad.equals(Obligatoriedad.Requerido)
					&& cell.getStringCellValue().isEmpty()) {
				estado = false;
				observacion = "Campo " + encabezado + " no debe ser nulo";
				break;
			}

			if (cell.getStringCellValue().isEmpty()) {
				cell = null;
				insertAddValues(propiedadColumnaExcel.getTipo(), cell, 0,
						insertValues);
				break;
			}

			if (!propiedadColumnaExcel.getRegex().isEmpty()) {
				estado = evaluarExpresionTexto(
						propiedadColumnaExcel.getRegex(), cell, encabezado);
				// break;
			}

			if (isFk) {
				listadoUnicoId.add(cell.getStringCellValue());
				mapaUnicoId.put(pos, cell.getStringCellValue());
			}

			if (isId) {
				listadoUnicoId.add(cell.getStringCellValue());
				mapaUnicoId.put(pos, cell.getStringCellValue());
			}

			insertAddValues(propiedadColumnaExcel.getTipo(), cell, 0,
					insertValues);
			// System.out.print(cell.getStringCellValue() + " ");
			break;
		case XSSFCell.CELL_TYPE_BLANK:
			if (obligatoriedad.equals(Obligatoriedad.Requerido)) {
				estado = false;
				observacion = "Campo " + encabezado + " no debe ser nulo";
			}
			insertAddValues(propiedadColumnaExcel.getTipo(), cell, 0,
					insertValues);
			// System.out.print(cell.getRawValue() + " ");
			break;
		case XSSFCell.CELL_TYPE_BOOLEAN:
			// System.out.print(cell.getBooleanCellValue() + " ");
			estado = false;
			observacion = "Campo " + encabezado + " celda no debe ser booleana";
			break;
		case XSSFCell.CELL_TYPE_ERROR:
			// System.out.print(cell.getErrorCellString().trim() + " ");
			estado = false;
			observacion = "Campo " + encabezado
					+ " existe un error en la celda";
			estado = false;
			break;
		case XSSFCell.CELL_TYPE_FORMULA:
			// System.out.print(cell.getRawValue().trim() + " ");
			estado = false;
			observacion = "Campo " + encabezado + " no debe ser una formula";
			break;
		case XSSFCell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				// System.out.print(cell.getDateCellValue() + " ");
				estado = false;
				observacion = "Campo " + encabezado
						+ " celda no debe ser de tipo fecha";
				break;
			}

			if (obligatoriedad.equals(Obligatoriedad.Requerido)
					&& cell.getRawValue().isEmpty()) {
				estado = false;
				observacion = "Campo " + encabezado + " no debe ser nulo";
				break;
			}

			if (!propiedadColumnaExcel.getRegex().isEmpty()) {
				switch (TipoDato.valueOf(propiedadColumnaExcel.getTipo())) {
				case Entero:
					estado = evaluarExpresionNumeroEntero(
							propiedadColumnaExcel.getRegex(), cell, encabezado);
					break;

				case Doble:
					estado = evaluarExpresionNumeroDouble(
							propiedadColumnaExcel.getRegex(), cell,
							propiedadColumnaExcel.getDecimalPlaces(),
							encabezado);
					break;

				default:
					estado = false;
					break;
				}

			}

			if (isFk) {
				listadoUnicoId.add(cell.getRawValue());
				mapaUnicoId.put(pos, cell.getRawValue());
			}

			if (isId) {
				listadoUnicoId.add(cell.getRawValue());
				mapaUnicoId.put(pos, cell.getRawValue());
			}

			insertAddValues(propiedadColumnaExcel.getTipo(), cell,
					propiedadColumnaExcel.getDecimalPlaces(), insertValues);
			break;

		default:
			System.out.print("No definida");
			estado = false;
			observacion = "Campo " + encabezado + " celda no definida";
			break;
		}
		return estado;
	}

}
