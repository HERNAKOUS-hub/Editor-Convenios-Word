package com.daw.cambiarword.Service;

import com.daw.cambiarword.modelaje.ItemRojo;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import org.springframework.stereotype.Service;
import jakarta.xml.bind.JAXBElement;

import java.io.*;
import java.util.*;

@Service
public class WordService {

    // Esta fábrica nos permite crear nuevos elementos XML (nuevos colores, textos, etc)
    private final org.docx4j.wml.ObjectFactory factory = new org.docx4j.wml.ObjectFactory();

    // ==========================================
    // FASE 1: EXTRACCIÓN (Encontrar huecos)
    // ==========================================
    public List<ItemRojo> extraerCadenasRojas(byte[] archivoBytes) throws Exception {
        List<ItemRojo> lista = new ArrayList<>();
        
        // 1. Cargar el array de bytes en un paquete completo de Word (WordprocessingMLPackage)
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new ByteArrayInputStream(archivoBytes));
        
        // 2. Extraer TODOS los párrafos del documento usando nuestra función recursiva.
        // No importa si están sueltos o escondidos en una tabla.
        List<P> parrafos = getAllParagraphs(wordMLPackage.getMainDocumentPart());
        int[] contador = { 1 };
        
        // 3. Procesarlos uno a uno
        procesarListaParrafosExtraccion(parrafos, lista, contador);
        
        return lista;
    }

    private void procesarListaParrafosExtraccion(List<P> parrafos, List<ItemRojo> lista, int[] contador) {
        for (P p : parrafos) {
            // Obtenemos todo el texto del párrafo para usarlo como "contexto visual" en la web
            String contextoParrafo = getParagraphText(p).trim();
            StringBuilder acumulado = new StringBuilder();
            boolean capturando = false;

            // Iteramos por todo el contenido del párrafo (Fragmentos 'R', Imágenes, etc)
            for (Object rObj : p.getContent()) {
                if (rObj instanceof R) {
                    R r = (R) rObj; // Si es un fragmento de texto (Run)...
                    
                    if (esObjetivo(r)) { 
                        // Si está en rojo o es un marcador ("NNNN"), guardamos su texto
                        acumulado.append(getRunText(r));
                        capturando = true;
                    } else {
                        // Si nos cruzamos con texto negro y estábamos capturando, es que el marcador terminó.
                        if (capturando) {
                            evaluarYAgregar(acumulado.toString(), contextoParrafo, lista, contador);
                            acumulado.setLength(0);
                            capturando = false;
                        }
                    }
                }
            }
            // Si el párrafo terminó y estábamos en medio de capturar texto, lo añadimos
            if (capturando) {
                evaluarYAgregar(acumulado.toString(), contextoParrafo, lista, contador);
            }
        }
    }

    private void evaluarYAgregar(String texto, String contexto, List<ItemRojo> lista, int[] contador) {
        String limpio = texto.trim();
        // Solo lo añadimos si pasa nuestro filtro (es corto y válido)
        if (esCampoValido(limpio)) {
            lista.add(new ItemRojo(contador[0]++, limpio, contexto));
        }
    }

    // ==========================================
    // FASE 2: REEMPLAZO (Aplicar cambios)
    // ==========================================
    public byte[] generarWordModificado(byte[] original, List<ItemRojo> cambios) throws Exception {
        // Volvemos a cargar el XML completo
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new ByteArrayInputStream(original));
        List<P> parrafos = getAllParagraphs(wordMLPackage.getMainDocumentPart());
        int[] contador = { 1 };
        
        // Buscamos los marcadores de nuevo y les aplicamos los cambios del usuario
        aplicarCambiosLista(parrafos, cambios, contador);
        
        // Convertimos el árbol XML modificado en un archivo de bytes para poder descargarlo
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        wordMLPackage.save(baos);
        return baos.toByteArray();
    }

    private void aplicarCambiosLista(List<P> parrafos, List<ItemRojo> cambios, int[] contador) {
        for (P p : parrafos) {
            List<R> grupo = new ArrayList<>();
            for (Object rObj : p.getContent()) {
                if (rObj instanceof R) {
                    R r = (R) rObj;
                    if (esObjetivo(r)) {
                        grupo.add(r);
                    } else {
                        if (!grupo.isEmpty()) {
                            evaluarYReemplazar(grupo, cambios, contador);
                            grupo.clear();
                        }
                    }
                }
            }
            if (!grupo.isEmpty()) {
                evaluarYReemplazar(grupo, cambios, contador);
            }
        }
    }

    private void evaluarYReemplazar(List<R> grupo, List<ItemRojo> cambios, int[] contador) {
        StringBuilder sb = new StringBuilder();
        for (R r : grupo) {
            sb.append(getRunText(r));
        }
        String limpio = sb.toString().trim();

        // Si era un marcador de verdad, lo reemplazamos
        if (esCampoValido(limpio)) {
            reemplazar(grupo, cambios, contador[0]++);
        }
    }

    private void reemplazar(List<R> grupo, List<ItemRojo> cambios, int pos) {
        // Buscamos en la lista de datos del usuario el texto que corresponde a esta posición
        ItemRojo item = cambios.stream().filter(i -> i.getPosicion() == pos).findFirst().orElse(null);

        if (item != null && item.getTextoNuevo() != null && !item.getTextoNuevo().trim().isEmpty()) {
            R primerRun = grupo.get(0);
            
            // 1. Cambiamos el texto
            setRunText(primerRun, item.getTextoNuevo());
            
            // 2. Modificamos las propiedades (RPr) para poner el texto de color NEGRO
            if (primerRun.getRPr() == null) {
                primerRun.setRPr(factory.createRPr());
            }
            Color color = factory.createColor();
            color.setVal("000000"); // Código HEX del color negro
            primerRun.getRPr().setColor(color);
            
            // 3. Si el texto original tenía marcador fluorescente, lo borramos
            if (primerRun.getRPr().getHighlight() != null) {
                primerRun.getRPr().setHighlight(null);
            }

            // 4. Si el marcador original ocupaba varios fragmentos de formato, borramos el resto para no duplicar texto
            for (int i = 1; i < grupo.size(); i++) {
                setRunText(grupo.get(i), "");
            }
        }
    }

    // ==========================================
    // HERRAMIENTAS Y NAVEGACIÓN DEL ÁRBOL XML
    // ==========================================

    /**
     * MAGIA PURA: Función recursiva. Como docx4j no te da los párrafos fácilmente si están
     * metidos en tablas, esto navega por todas las "ramas" del árbol XML y extrae cualquier Párrafo (P) que vea.
     */
    private List<P> getAllParagraphs(Object obj) {
        List<P> result = new ArrayList<>();
        if (obj instanceof P) {
            result.add((P) obj);
        } else if (obj instanceof ContentAccessor) { // Si es algo que puede contener hijos (ej. Tablas, Celdas)
            for (Object child : ((ContentAccessor) obj).getContent()) {
                result.addAll(getAllParagraphs(child));
            }
        } else if (obj instanceof JAXBElement) { // Nodos XML genéricos
            result.addAll(getAllParagraphs(((JAXBElement<?>) obj).getValue()));
        }
        return result;
    }

    /** Extrae todo el texto limpio de un Párrafo juntando sus pedazos */
    private String getParagraphText(P p) {
        StringBuilder sb = new StringBuilder();
        for (Object rObj : p.getContent()) {
            if (rObj instanceof R) {
                sb.append(getRunText((R) rObj));
            }
        }
        return sb.toString();
    }

    /** Extrae el texto real (Text) que hay dentro de un fragmento de formato (R) */
    private String getRunText(R r) {
        StringBuilder sb = new StringBuilder();
        for (Object obj : r.getContent()) {
            if (obj instanceof JAXBElement) {
                Object val = ((JAXBElement<?>) obj).getValue();
                if (val instanceof Text) {
                    sb.append(((Text) val).getValue());
                }
            } else if (obj instanceof Text) {
                sb.append(((Text) obj).getValue());
            }
        }
        return sb.toString();
    }

    /** Elimina el texto viejo de un XML e inyecta uno nuevo con espacios preservados */
    private void setRunText(R r, String nuevoTexto) {
        // Borramos el XML de texto antiguo
        r.getContent().removeIf(obj -> 
            (obj instanceof JAXBElement && ((JAXBElement<?>) obj).getValue() instanceof Text) || 
            obj instanceof Text
        );
        
        // Creamos la etiqueta <w:t> del nuevo texto
        if (nuevoTexto != null && !nuevoTexto.isEmpty()) {
            Text t = factory.createText();
            t.setValue(nuevoTexto);
            t.setSpace("preserve"); // Evita que Word se coma los espacios iniciales
            
            // ¡AQUÍ ESTABA EL ERROR! Es createRT, no createRText
            JAXBElement<Text> jaxbText = factory.createRT(t); 
            
            r.getContent().add(jaxbText);
        }
    }

    // ==========================================
    // LOS FILTROS DE NEGOCIO
    // ==========================================

    /** Comprueba si un fragmento de Word está en color distinto al negro o tiene NNN/xxx */
    private boolean esObjetivo(R r) {
        RPr rpr = r.getRPr(); // rpr = Propiedades del Run (Color, tamaño, etc)
        if (rpr != null) {
            Color color = rpr.getColor();
            if (color != null && color.getVal() != null) {
                String val = color.getVal();
                // "000000" es negro. "auto" es el negro por defecto de Word.
                if (!val.equals("000000") && !val.equalsIgnoreCase("auto")) {
                    return true;
                }
            }
            // Comprobamos también si es un marcador fluorescente
            Highlight hl = rpr.getHighlight();
            if (hl != null && hl.getVal() != null && !hl.getVal().equals("none")) {
                return true;
            }
        }

        // Plan B: Buscar por el contenido literal del texto
        String texto = getRunText(r);
        if (texto != null) {
            String t = texto.toLowerCase();
            if (t.contains("nnn") || t.contains("xxx")) {
                return true;
            }
        }
        return false;
    }

    /** Filtro que ignora las leyes enteras rojas (muy largas) pero atrapa huecos variables */
    private boolean esCampoValido(String texto) {
        if (texto == null || texto.isEmpty()) return false;
        if (!texto.matches(".*[a-zA-Z0-9].*")) return false; // Evitar capturar espacios rojos en blanco
        
        // Si tiene más de 65 caracteres, asumimos que es una cláusula legal y no un hueco.
        if (texto.length() > 65) return false;
        
        return true;
    }
}
