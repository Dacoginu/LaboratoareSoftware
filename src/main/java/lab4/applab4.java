package lab4;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;

public class applab4 {
    public static void main(String[] args) {
        HashMap<String, Integer> varste = new HashMap<>();
        varste.put("Ioan", 21);
        varste.put("Maria", 22);
        varste.put("Victor", 20);
        varste.put("Simina", 20);
        varste.put("Marius", 21);
        varste.put("Mihai", 21);
        varste.put("Daniela", 23);

        Map<String, String> adrese = Map.of("Ioan", "Sibiu", "Maria", "Bucuresti", "Victor",
                "Cluj", "Simina", "Alba-Iulia", "Marius", "Medias", "Mihai", "Cisnadie", "Daniela", "Sibiu");


        System.out.println("a)");
        System.out.println("Continut initial varste:");
        System.out.println(varste);

        varste.put("Vlad", 19);
        varste.put("Iulia", 19);

        System.out.println("\n");
        System.out.println(varste);

        System.out.println();
        HashMap<String, Tanar> tineri = new HashMap<>();
        for (String nume : varste.keySet()) {
            int ani = varste.get(nume);

            String addr = adrese.get(nume);
            Tanar t = new Tanar(nume, ani, addr);
            tineri.put(nume, t);
        }

            System.out.println("d)");
            System.out.println("Continutul instantei tineri:");
            for (Object entry : tineri.values()) {
                System.out.println(entry.toString());

            }

        }
    }
