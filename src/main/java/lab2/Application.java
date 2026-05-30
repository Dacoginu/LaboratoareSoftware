package lab2;
import java.util.*;

public class Application {
    public static void main(String[] args) {

        List<Integer> x = new ArrayList<>();
        List<Integer> y = new ArrayList<>();
        List<Integer> xPlusY = new ArrayList<>();
        Set<Integer> zSet = new TreeSet<>();
        List<Integer> xMinusY = new ArrayList<>();
        int p = 4;
        List<Integer> xPlusYLimitedByP = new ArrayList<>();

        Random rand = new Random();

        for (int i = 0; i < 5; i++) {
            x.add(rand.nextInt(11));
        }

        for (int i = 0; i < 7; i++) {
            y.add(rand.nextInt(11));
        }

        Collections.sort(x);
        Collections.sort(y);

        xPlusY.addAll(x);
        xPlusY.addAll(y);
        Collections.sort(xPlusY);

        for (Integer val : x) {
            if(y.contains(val))
                zSet.add(val);
        }

        for (Integer val : x) {
            if(!y.contains(val))
                xMinusY.add(val);
        }

        for (Integer val : xPlusY) {
            if(val <= p)
                xPlusYLimitedByP.add(val);
        }

        System.out.println("Lista x: " + x);
        System.out.println("Lista y: " + y);
        System.out.println("a) xPlusY (reuniune): " + xPlusY);
        System.out.println("b) zSet (intersectie): " + zSet);
        System.out.println("c) xMinusY (diferenta x - y): " + xMinusY);
        System.out.println("d) xPlusYLimitedByP (valori <= " + p + "): " + xPlusYLimitedByP);
    }
}