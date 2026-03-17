package ro.ulbs.proiectaresoftware.lab1;

public class Application {
public static void main(String[] args) {
String s = "";

for (char i = 'a'; i <= 'z'; i++) {
s += i;
 }

System.out.println(s.toLowerCase());
 System.out.println(s.toUpperCase());
 String[] array = new String[s.length()];
 int k = 0;
 String segment = "";

 for (int i = 0; i < s.length(); i++) {
 char c = s.charAt(i);
 if ("aeiou".indexOf(c) != -1) {
if (!segment.equals("")) {
 array[k] = segment;
 k++;
 }
 segment = "" + c;
 } else {
 if (!segment.equals("")) {
 segment += c;
 }
 }
 }


 if (!segment.equals("")) {
 array[k] = segment;
 k++;
 }


for (int i = 0; i < k; i++) {
System.out.println("array[" + i + "] =  " + array[i]);} }
}