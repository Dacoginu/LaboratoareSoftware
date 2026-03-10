package lab2;
import java.util.*;
public class Application {
    public static void main(String[] args) {


        List<Integer> x = new ArrayList();
        List<Integer> y = new ArrayList();
        List<Integer> xPlusY = new ArrayList(); //a
        Set<Integer> zSet = new TreeSet(); //b
        List<Integer> xMinusY = new ArrayList();//c
        int p = 4;
        List<Integer> xPlusYLimitedByP = new ArrayList(); //d
    Random rand=new Random();

      for (int i = 0; i < 5; i++) {
          x.add(rand.nextInt(11));
      }
      for (int i = 0; i < 7; i++) {
          y.add(rand.nextInt(11));
          Collections.sort(x);
          Collections.sort(y);
          xPlusY.addAll(x);
          xPlusY.addAll(y);
          Collections.sort(xPlusY);
          for (Integer val : x)
          {
              if(y.contains(val))
                  zSet.add(val);
          }
          for (Integer val : x)
          {
              if(!y.contains(val))
                  xMinusY.add(val);
          }
          for (Integer val : xPlusY)
          {
              if(val<p)
                  xPlusYLimitedByP.add(val);
          }

      }
    }
}
