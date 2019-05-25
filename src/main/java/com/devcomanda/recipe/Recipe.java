package com.devcomanda.recipe;

import java.util.*;

public class Recipe {

    private String name;
    private Map<String, String> ingredients = new HashMap<>();
    private List<String> items = new ArrayList<>();
    private String urlToImage;

    public Recipe(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }

    public Iterator<Map.Entry<String, String>> getIngredients() {
        return ingredients.entrySet().iterator();
    }

    public void addIngredient(String name, String amount) {
        this.ingredients.put(name, amount);
    }

    public Iterator<String> getItems() {
        return items.iterator();
    }

    public void addItem(String item) {
        this.items.add(item);
    }

    public String getUrlToImage() {
        return urlToImage;
    }

    public void setUrlToImage(String urlToImage) {
        this.urlToImage = urlToImage;
    }
}
