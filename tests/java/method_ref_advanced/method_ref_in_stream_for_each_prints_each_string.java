// vybe-test: java/method_ref_advanced/method_ref_in_stream_for_each_prints_each_string
// origin: languages/java/tests/java/test_method_ref_advanced.rs

public class Main {
    public static void main(String[] args) {
java.util.Arrays.asList("x", "y").stream().forEach(System.out::println);
    }
}

