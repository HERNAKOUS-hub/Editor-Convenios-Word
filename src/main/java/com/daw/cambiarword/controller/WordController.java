package com.daw.cambiarword.controller;

import com.daw.cambiarword.modelaje.*;
import com.daw.cambiarword.Service.WordService;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import jakarta.servlet.http.*;
import java.util.List;

@Controller
public class WordController {

    // Inyectamos el servicio que hace la magia con el Word
    private final WordService wordService;

    public WordController(WordService wordService) {
        this.wordService = wordService;
    }

    // ==========================================
    // PANTALLA INICIAL (Subir archivo)
    // ==========================================
    @GetMapping("/")
    public String index() {
        return "upload"; // Muestra la vista upload.html
    }

    // ==========================================
    // PROCESAR EL ARCHIVO SUBIDO
    // ==========================================
    // Fíjate en el "throws Exception" porque docx4j puede lanzar varios tipos de errores al leer el XML
    @PostMapping("/procesar-word")
    public String procesar(@RequestParam("archivo") MultipartFile archivo, HttpSession session) throws Exception {
        // 1. Obtenemos los "bytes" puros del archivo que ha subido el usuario
        byte[] bytes = archivo.getBytes();
        
        // 2. Guardamos el archivo original en la Sesión (la memoria temporal del usuario) 
        // para tenerlo disponible cuando el usuario termine de rellenar el formulario.
        session.setAttribute("archivoOriginal", bytes);

        // 3. Llamamos al servicio para que nos busque todos los huecos a rellenar
        List<ItemRojo> lista = wordService.extraerCadenasRojas(bytes);
        
        // 4. Guardamos esa lista en la sesión dentro de un "FormularioReemplazo"
        session.setAttribute("formulario", new FormularioReemplazo(lista));

        // 5. Redirigimos a la pantalla de edición
        return "redirect:/editar";
    }

    // ==========================================
    // PANTALLA DE EDICIÓN
    // ==========================================
    @GetMapping("/editar")
    public String editar(HttpSession session, Model model) {
        // 1. Recuperamos el formulario de la sesión
        FormularioReemplazo form = (FormularioReemplazo) session.getAttribute("formulario");
        
        // Si el usuario llega aquí sin haber subido un archivo, le devolveímosa al inicio
        if (form == null) {
            return "redirect:/";
        }

        // 2. Le pasamos el formulario a la vista editar.html para que pinte las cajas de texto
        model.addAttribute("formulario", form);
        return "editar";
    }

    // ==========================================
    // APLICAR CAMBIOS Y DESCARGAR
    // ==========================================
    @PostMapping("/aplicar-cambios")
    public void finalizar(@ModelAttribute("formulario") FormularioReemplazo form, HttpSession session,
            HttpServletResponse response) throws Exception {
                
        // 1. Recuperamos el Word original intacto que guardamos al principio
        byte[] original = (byte[]) session.getAttribute("archivoOriginal");
        if (original == null) {
            response.sendRedirect("/");
            return;
        }

        // 2. Le decimos al servicio: "Toma el Word original y la lista de textos que ha escrito el usuario, y dame un Word nuevo"
        byte[] modificado = wordService.generarWordModificado(original, form.getItems());

        // 3. Configuramos la respuesta HTTP para obligar al navegador a "descargar" un archivo en vez de mostrar una web
        response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        response.setHeader("Content-Disposition", "attachment; filename=Convenio_Editado.docx");
        
        // 4. Escribimos los bytes del nuevo Word en la salida hacia el navegador
        response.getOutputStream().write(modificado);
        response.getOutputStream().flush();
    }
}
