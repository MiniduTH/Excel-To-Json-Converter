package excelConvertor;

import java.io.BufferedWriter;
import java.io.FileWriter;

public class WriteJson {
    private static String data;
    private static String outputLocation;

    WriteJson(String str,String output){
        data = str;
        outputLocation = output;

    }
    public static void main(String[] args) {
    }

    public void fileWrite(){
        try {
            BufferedWriter writer = new BufferedWriter(new FileWriter(outputLocation+"\\Output.json"));
            writer.write(data);
            writer.close();
            System.out.println("written successfully");

        } catch (Exception e) {
            e.printStackTrace();
        }

    }


}
