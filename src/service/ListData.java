package service;

import java.io.File;
import java.util.ArrayList;

public class ListData {
    public ArrayList<String> ReadDirectory() {
        String path = "C:\\Users\\768901\\Downloads";

        ArrayList<String> files = new ArrayList<String>();
        File folder = new File(path);
        File[] listOfFiles = folder.listFiles();

        for (int i = 0; i < listOfFiles.length; i++) {
            if (listOfFiles[i].isFile()) {

                if(listOfFiles[i].getName().endsWith(".docx")||listOfFiles[i].getName().endsWith("doc")){
                    files.add(listOfFiles[i].getName());
                }

            }
        }
        return files;
    }
}
