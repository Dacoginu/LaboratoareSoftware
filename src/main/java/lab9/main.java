package lab9;
import java.util.*;

public class main {
    public static void main(String[] args) {
        Random rand=new Random();
        List<Integer> lista = rand.ints(10, 5, 26)
                .boxed()
                .toList();
        System.out.println("Lista initiala: " + lista);


//a
        int suma = lista.stream()
                .mapToInt(Integer::intValue)
                .sum();
        System.out.println("Suma: " + suma);

//b
        int max = lista.stream().mapToInt(Integer::intValue).max().orElse(0);
        int min = lista.stream().mapToInt(Integer::intValue).min().orElse(0);
        System.out.println("Max: " + max);
        System.out.println("Min: " + min);

//c
List<Integer> filtrata = lista.stream()
        .filter(x -> x >= 10 && x <= 20)
        .toList();
        System.out.println("Lista filtrata [10..20]: " + filtrata);
//d
        List<Double> listaDouble = lista.stream()
                .map(Integer::doubleValue)
                .toList();
        System.out.println("Lista Double: " + listaDouble);
//e
        boolean exista12 = lista.stream().anyMatch(x -> x == 12);
        System.out.println("Exista 12? " + exista12);

        String text = "Acesta este un program scris in java pentru expresii lambda";


        List<String> cuvinte = Arrays.asList(text.split(" "));
        System.out.println("Lista initiala: " + cuvinte);
//a
        List<String> filtratat = cuvinte.stream()
                .filter(c -> c.length() >= 5)
                .toList();

        long count = filtratat.size();

        System.out.println("Lista filtrata (>=5 caractere): " + filtratat);
        System.out.println("Numar cuvinte: " + count);
// b)
        List<String> sortata = filtratat.stream()
                .sorted()
                .toList();

        System.out.println("Lista sortata: " + sortata);

// c)
        Optional<String> cuvantP = cuvinte.stream()
                .filter(c -> c.startsWith("p"))
                .findFirst();

        cuvantP.ifPresentOrElse(
                c -> System.out.println("Cuvant care incepe cu 'p': " + c),
                () -> System.out.println("Nu exista cuvant care incepe cu 'p'")
        );


    }

}
