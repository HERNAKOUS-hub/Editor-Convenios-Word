package com.daw.cambiarword.Service;

import com.daw.cambiarword.modelaje.ItemRojo;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;
import java.io.*;
import java.util.*;

@Service
public class WordService {

    public List<ItemRojo> extraerCadenasRojas(byte[] archivoBytes) throws IOException {
        List<ItemRojo> lista = new ArrayList<>();
        try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(archivoBytes))) {
            int[] contador = { 1 };
            procesarCuerpo(doc, lista, contador);
        }
        return lista;
    }

    private void procesarCuerpo(IBody cuerpo, List<ItemRojo> lista, int[] contador) {
        procesarListaParrafosExtraccion(cuerpo.getParagraphs(), lista, contador);
        for (XWPFTable tabla : cuerpo.getTables()) {
            procesarTablaExtraccion(tabla, lista, contador);
        }
    }

    private void procesarTablaExtraccion(XWPFTable tabla, List<ItemRojo> lista, int[] contador) {
        for (XWPFTableRow fila : tabla.getRows()) {
            for (XWPFTableCell celda : fila.getTableCells()) {
                procesarListaParrafosExtraccion(celda.getParagraphs(), lista, contador);
                for (XWPFTable tablaAnidada : celda.getTables()) {
                    procesarTablaExtraccion(tablaAnidada, lista, contador);
                }
            }
        }
    }

    private void procesarListaParrafosExtraccion(List<XWPFParagraph> parrafos, List<ItemRojo> lista, int[] contador) {
        for (XWPFParagraph p : parrafos) {
            String contextoParrafo = p.getText().trim();
            StringBuilder acumulado = new StringBuilder();
            boolean capturando = false;

            for (XWPFRun r : p.getRuns()) {
                if (esObjetivo(r)) {
                    if (r.getText(0) != null) {
                        acumulado.append(r.getText(0));
                    }
                    capturando = true;
                } else {
                    if (capturando) {
                        evaluarYAgregar(acumulado.toString(), contextoParrafo, lista, contador);
                        acumulado.setLength(0);
                        capturando = false;
                    }
                }
            }
            if (capturando) {
                evaluarYAgregar(acumulado.toString(), contextoParrafo, lista, contador);
            }
        }
    }

    private void evaluarYAgregar(String texto, String contexto, List<ItemRojo> lista, int[] contador) {
        String limpio = texto.trim();
        if (esCampoValido(limpio)) {
            lista.add(new ItemRojo(contador[0]++, limpio, contexto));
        }
    }

    // =========================================================================
    // LÓGICA DE REEMPLAZO
    // =========================================================================

    public byte[] generarWordModificado(byte[] original, List<ItemRojo> cambios) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(new ByteArrayInputStream(original));
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            int[] contador = { 1 };
            aplicarCuerpo(doc, cambios, contador);
            doc.write(baos);
            return baos.toByteArray();
        }
    }

    private void aplicarCuerpo(IBody cuerpo, List<ItemRojo> cambios, int[] contador) {
        aplicarCambiosLista(cuerpo.getParagraphs(), cambios, contador);
        for (XWPFTable tabla : cuerpo.getTables()) {
            aplicarTabla(tabla, cambios, contador);
        }
    }

    private void aplicarTabla(XWPFTable tabla, List<ItemRojo> cambios, int[] contador) {
        for (XWPFTableRow fila : tabla.getRows()) {
            for (XWPFTableCell celda : fila.getTableCells()) {
                aplicarCambiosLista(celda.getParagraphs(), cambios, contador);
                for (XWPFTable tablaAnidada : celda.getTables()) {
                    aplicarTabla(tablaAnidada, cambios, contador);
                }
            }
        }
    }

    private void aplicarCambiosLista(List<XWPFParagraph> parrafos, List<ItemRojo> cambios, int[] contador) {
        for (XWPFParagraph p : parrafos) {
            List<XWPFRun> grupo = new ArrayList<>();
            for (XWPFRun r : p.getRuns()) {
                if (esObjetivo(r)) {
                    grupo.add(r);
                } else {
                    if (!grupo.isEmpty()) {
                        evaluarYReemplazar(grupo, cambios, contador);
                        grupo.clear();
                    }
                }
            }
            if (!grupo.isEmpty()) {
                evaluarYReemplazar(grupo, cambios, contador);
            }
        }
    }

    private void evaluarYReemplazar(List<XWPFRun> grupo, List<ItemRojo> cambios, int[] contador) {
        StringBuilder sb = new StringBuilder();
        for (XWPFRun r : grupo) {
            if (r.getText(0) != null)
                sb.append(r.getText(0));
        }

        String limpio = sb.toString().trim();

        if (esCampoValido(limpio)) {
            reemplazar(grupo, cambios, contador[0]++);
        }
    }

    private void reemplazar(List<XWPFRun> grupo, List<ItemRojo> cambios, int pos) {
        ItemRojo item = cambios.stream().filter(i -> i.getPosicion() == pos).findFirst().orElse(null);

        if (item != null && item.getTextoNuevo() != null && !item.getTextoNuevo().trim().isEmpty()) {
            grupo.get(0).setText(item.getTextoNuevo(), 0);
            grupo.get(0).setColor("000000"); // Ponemos el texto final en negro

            for (int i = 1; i < grupo.size(); i++) {
                grupo.get(i).setText("", 0);
            }
        }
    }

    // =========================================================================
    // EL BISTURÍ: Filtros de texto exactos para tu documento
    // =========================================================================

    private boolean esObjetivo(XWPFRun r) {
        // 1. Buscamos por color (Cualquier cosa que NO sea negro o automático)
        String color = r.getColor();
        if (color != null && !color.equals("000000") && !color.equalsIgnoreCase("auto")) {
            return true;
        }

        // 2. Por si el color falla, buscamos los patrones de tus marcadores
        String texto = r.getText(0);
        if (texto != null) {
            String t = texto.toLowerCase();
            if (t.contains("nnn") || t.contains("xxx")) {
                return true;
            }
        }
        return false;
    }

    private boolean esCampoValido(String texto) {
        if (texto == null || texto.isEmpty())
            return false;

        // Evitamos capturar comas sueltas o espacios que Word haya puesto en rojo por
        // error
        if (!texto.matches(".*[a-zA-Z0-9].*"))
            return false;

        // MAGIA: Filtro de longitud.
        // Tus marcadores como "xxxxxxxxxxxxxxxxxx" o "N de NNNNNNNN de NNNNN" son
        // cortos.
        // Un párrafo legal entero tiene cientos de letras.
        // Si tiene más de 65 caracteres, LO IGNORAMOS, no es un hueco a rellenar.
        if (texto.length() > 65) {
            return false;
        }

        return true;
    }
}