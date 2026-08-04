// vybe-test: java/generic_core/wildcard_list_accepts_string_elements
// origin: languages/java/tests/java/test_generic_core.rs

public class Main {
static class Printers {
            static void printAll(java.util.List<?> items) {
                for (Object o : items) System.out.println(o);
            }
        }
    public static void main(String[] args) {
java.util.List<String> words = java.util.Arrays.asList("a", "b"); Printers.printAll(words);
    }
}

