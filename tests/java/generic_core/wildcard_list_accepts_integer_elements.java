// vybe-test: java/generic_core/wildcard_list_accepts_integer_elements
// origin: languages/java/tests/java/test_generic_core.rs

public class Main {
static class Printers {
            static void printAll(java.util.List<?> items) {
                for (Object o : items) System.out.println(o);
            }
        }
    public static void main(String[] args) {
java.util.List<Integer> nums = java.util.Arrays.asList(1, 2); Printers.printAll(nums);
    }
}

