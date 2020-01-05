package com.pranshu.exps.util;

import java.util.ArrayList;
import java.util.List;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class InputFileReader {
	public static List<Object> readFile(String fileName) {
		List<Object> records = new ArrayList<>();
		try (Stream<String> stream = Files.lines(Paths.get(fileName))) {
			records = stream.collect(Collectors.toList());
		} catch (IOException e) {
			e.printStackTrace();
		}
		return records;
	}
}
