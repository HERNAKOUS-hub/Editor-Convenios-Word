package com.daw.cambiarword.modelaje;

public class ItemRojo {

    private int posicion;
    private String textoOriginal;
    private String textoNuevo;
    private String contexto; // <-- NUEVA VARIABLE

    public ItemRojo() {
    }

    public ItemRojo(int posicion, String textoOriginal, String contexto) {
        this.posicion = posicion;
        this.textoOriginal = textoOriginal;
        this.contexto = contexto;
    }

    public int getPosicion() {
        return posicion;
    }

    public void setPosicion(int posicion) {
        this.posicion = posicion;
    }

    public String getTextoOriginal() {
        return textoOriginal;
    }

    public void setTextoOriginal(String textoOriginal) {
        this.textoOriginal = textoOriginal;
    }

    public String getTextoNuevo() {
        return textoNuevo;
    }

    public void setTextoNuevo(String textoNuevo) {
        this.textoNuevo = textoNuevo;
    }

    public String getContexto() {
        return contexto;
    }

    public void setContexto(String contexto) {
        this.contexto = contexto;
    }
}