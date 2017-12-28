package com.exco.cargaExcel.enums;

/**
 * Enumerador con las opciones (Requerido, Opcional) y su acrónimo
 * 
 * @author EXCO
 *
 */
public enum Obligatoriedad {
	
	Requerido("R"),Opcional("O");

	public String estado;
	
	private Obligatoriedad(String estado) {
		this.estado = estado;
	}
	public String mostrar() {
		return estado;
	}
}
