package com.emoisei.springcourse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;

@Component
public class MusicPlayer {

  private ClassicalMusic classicalMusic;
  private RockMusic rockMusic;
@Autowired
    public MusicPlayer(@Qualifier("rockMusic") RockMusic rockMusic,
                       @Qualifier("classicalMusic") ClassicalMusic classicalMusic) {
        this.rockMusic = rockMusic;
        this.classicalMusic = classicalMusic;
    }

    public void playMusic(MusicGenre genre) {
        Random random = new Random();

        int randomNumber = random.nextInt(3);

        if(genre == MusicGenre.CLASSICAL){
            System.out.println(classicalMusic.getSongs().get(randomNumber));
        }
        else{
            System.out.println(rockMusic.getSongs().get(randomNumber));
        }



    }
}
