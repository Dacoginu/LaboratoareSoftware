package lab3;

import java.io.*;
import java.util.*;

public class Application {
    public static void main(String[] args) {
        List<String> linii = new ArrayList<>();


        try (BufferedReader br = new BufferedReader(new FileReader("in.txt"))) {
            String linie;
            while ((linie = br.readLine()) != null) {

                String[] bucati = linie.split("\n");
                for (String b : bucati) {
                    linii.add(b);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        List<String> reza = new ArrayList<>();
        System.out.println("a");
        for (String linie : linii) {
            String nou = linie + "\n";
            reza.add(nou);
            System.out.print(nou);
        }
        List<String> rezb = new ArrayList<>();
        System.out.println("b");
        for (String linie : linii) {
            String nou = linie.replace(".", ".\n");
            rezb.add(nou);
            System.out.print(nou);
        }
        try (BufferedWriter bw = new BufferedWriter(new FileWriter("out.txt"))) {
            bw.write(" a\n");
            for (String linie : reza) {
                bw.write(linie);
            }

            bw.write("\n b\n");
            for (String linie : rezb) {
                bw.write(linie + "\n");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("\nDatele au fost salvate in out.txt");
    }
}