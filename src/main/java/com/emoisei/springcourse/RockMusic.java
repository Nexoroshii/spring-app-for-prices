package com.emoisei.springcourse;

import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;

@Component
public class RockMusic implements  Music{

    private List<String> songs = new ArrayList<>();
    {
        songs.add("rock1");
        songs.add("rock2");
        songs.add("rock3");
    }


    @Override
    public List<String> getSongs() {
        return songs;
    }
}
