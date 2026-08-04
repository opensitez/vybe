// vybe-test: java/lambda_core/method_reference_system_out_println_via_consumer
// origin: languages/java/tests/java/test_lambda_core.rs

public class Main {
    public static void main(String[] args) {
java.util.function.Consumer<String> log = System.out::println; log.accept("logged");
    }
}

