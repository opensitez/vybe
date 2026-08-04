// vybe-test: java/method_ref_advanced/method_ref_list_for_each_with_println_reference
// origin: languages/java/tests/java/test_method_ref_advanced.rs

public class Main {
    public static void main(String[] args) {
java.util.Arrays.asList(1, 2, 3).forEach(System.out::println);
    }
}

