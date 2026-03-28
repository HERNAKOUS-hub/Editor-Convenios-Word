package com.daw.cambiarword.controller;

import com.daw.cambiarword.modelaje.*;
import com.daw.cambiarword.Service.WordService;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import jakarta.servlet.http.*;
import java.io.IOException;
import java.util.List;

@Controller
public class WordController {

    private final WordService wordService;

    public WordController(WordService wordService) {
        this.wordService = wordService;
    }

    @GetMapping("/")
    public String index() {
        return "upload";
    }

    @PostMapping("/procesar-word")
    public String procesar(@RequestParam("archivo") MultipartFile archivo, HttpSession session) throws IOException {
        byte[] bytes = archivo.getBytes();
        session.setAttribute("archivoOriginal", bytes);

        List<ItemRojo> lista = wordService.extraerCadenasRojas(bytes);
        session.setAttribute("formulario", new FormularioReemplazo(lista));

        return "redirect:/editar";
    }

    @GetMapping("/editar")
    public String editar(HttpSession session, Model model) {
        FormularioReemplazo form = (FormularioReemplazo) session.getAttribute("formulario");
        if (form == null)
            return "redirect:/";

        model.addAttribute("formulario", form);
        return "editar";
    }

    @PostMapping("/aplicar-cambios")
    public void finalizar(@ModelAttribute("formulario") FormularioReemplazo form, HttpSession session,
            HttpServletResponse response) throws IOException {
        byte[] original = (byte[]) session.getAttribute("archivoOriginal");
        if (original == null) {
            response.sendRedirect("/");
            return;
        }

        byte[] modificado = wordService.generarWordModificado(original, form.getItems());

        response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        response.setHeader("Content-Disposition", "attachment; filename=Convenio_Editado.docx");
        response.getOutputStream().write(modificado);
        response.getOutputStream().flush();
    }
}