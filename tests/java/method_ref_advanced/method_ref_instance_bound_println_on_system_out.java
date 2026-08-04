// vybe-test: java/method_ref_advanced/method_ref_instance_bound_println_on_system_out
// origin: languages/java/tests/java/test_method_ref_advanced.rs

public class Main {
    public static void main(String[] args) {
java.util.function.Consumer<String> log = System.out::println; log.accept("bound");
    }
}

