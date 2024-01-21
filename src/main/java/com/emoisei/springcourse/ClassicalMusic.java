package com.emoisei.springcourse;

import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;

@Component
public class ClassicalMusic implements Music {
 private List<String> songs = new ArrayList<>();
    {
        songs.add("classical1");
        songs.add("classical2");
        songs.add("classical3");
    }
    @Override
    public List<String> getSongs() {
        return songs;
    }


}
