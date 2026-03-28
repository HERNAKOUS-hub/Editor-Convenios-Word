package com.daw.cambiarword.modelaje;

import java.util.ArrayList;
import java.util.List;

public class FormularioReemplazo {

    private List<ItemRojo> items = new ArrayList<>();

    public FormularioReemplazo() {
    }

    public FormularioReemplazo(List<ItemRojo> items) {
        this.items = items;
    }

    public List<ItemRojo> getItems() {
        return items;
    }

    public void setItems(List<ItemRojo> items) {
        this.items = items;
    }
}