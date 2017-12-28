package com.exco.cargaExcel.model;

/**
 * Clase con las propiedades de las columnas de Excel mapeadas desde un archivo de configuración "*.yml"
 * 
 * @author EXCO
 *
 */
public class PropiedadColumnaExcel {
	
	public String nombre;
	public String tipo;
	public String obligatoriedad;
	public String regex;
	public int dependencia;
	public String nativeQuery;
	public int longitud;
	public int decimalPlaces;
	public double minValue;
	public double maxValue;
		
	public String getNombre() {
		return nombre;
	}
	public void setNombre(String nombre) {
		this.nombre = nombre;
	}
	public String getTipo() {
		return tipo;
	}
	public void setTipo(String tipo) {
		this.tipo = tipo;
	}
	public String getObligatoriedad() {
		return obligatoriedad;
	}
	public void setObligatoriedad(String obligatoriedad) {
		this.obligatoriedad = obligatoriedad;
	}
	public String getRegex() {
		return regex;
	}
	public void setRegex(String regex) {
		this.regex = regex;
	}
	public int getLongitud() {
		return longitud;
	}
	public void setLongitud(int longitud) {
		this.longitud = longitud;
	}
	public double getMinValue() {
		return minValue;
	}
	public void setMinValue(double minValue) {
		this.minValue = minValue;
	}
	public double getMaxValue() {
		return maxValue;
	}
	public void setMaxValue(double maxValue) {
		this.maxValue = maxValue;
	}
	public int getDependencia() {
		return dependencia;
	}
	public void setDependencia(int dependencia) {
		this.dependencia = dependencia;
	}
	public String getNativeQuery() {
		return nativeQuery;
	}
	public void setNativeQuery(String nativeQuery) {
		this.nativeQuery = nativeQuery;
	}
	public int getDecimalPlaces() {
		return decimalPlaces;
	}
	public void setDecimalPlaces(int decimalPlaces) {
		this.decimalPlaces = decimalPlaces;
	}
	
	
}
