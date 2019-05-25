package com.devcomanda.user;

public class User {
    private String name;
    private int experience;

    public User(String name, int experience) {
        this.name = name;
        this.experience = experience;
    }

    public String getName() {
        return name;
    }

    public int getExperience() {
        return experience;
    }
}
