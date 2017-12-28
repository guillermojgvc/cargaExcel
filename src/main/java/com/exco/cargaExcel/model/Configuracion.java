package com.exco.cargaExcel.model;

import java.io.FileReader;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import com.esotericsoftware.yamlbeans.YamlReader;

public class Configuracion {

	/**
	 * Metodo para leer el archivo de propiedades de cada columna del archivo de
	 * Excel desde un archivo yml
	 * 
	 * @param archivo
	 *            String con la ubicación y nombre del archivo
	 * @return Lista de PropiedadColumnaExcel.
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public List<PropiedadColumnaExcel> LeerArchivoPropiedades(String archivo) {
		YamlReader reader;
		List<PropiedadColumnaExcel> excelProperties = new ArrayList<PropiedadColumnaExcel>();
		try {
			// reader = new YamlReader(new
			// FileReader(Configuracion.class.getClass().getResource(archivo).getFile()));
			reader = new YamlReader(new FileReader(archivo));
			Object object;
			object = reader.read();
			Map map = (Map) object;
			excelProperties = (ArrayList<PropiedadColumnaExcel>) map
					.get("Columnas");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return excelProperties;
	}

	/**
	 * Metodo para leer el archivo de propiedades y obtener el native query
	 * desde un archivo yml para la inserción de datos
	 * 
	 * @param archivo
	 *            String con la ubicación y nombre del archivo
	 * @return InsertTemplate objeto con el native query.
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public InsertTemplate LeerArchivoPropiedadesInsertTemplate(String archivo) {
		YamlReader reader;
		List<InsertTemplate> insertTemplate = new ArrayList<InsertTemplate>();
		try {
			// reader = new YamlReader(new
			// FileReader(Configuracion.class.getClass().getResource(archivo).getFile()));
			reader = new YamlReader(new FileReader(archivo));
			Object object;
			object = reader.read();
			Map map = (Map) object;
			insertTemplate = (ArrayList<InsertTemplate>) map
					.get("InsertTemplate");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return insertTemplate.get(0);
	}
}
