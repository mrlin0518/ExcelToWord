package com.example.demo;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.stream.IntStream;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java_cup.internal_error;

@SpringBootApplication
public class ExcelToWordApplication {

	public static void main(String[] args) {
		SpringApplication.run(ExcelToWordApplication.class, args);
		ExcelToWord excelToWord = new ExcelToWord(); 
		try {
			excelToWord.excelToWordStart();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
