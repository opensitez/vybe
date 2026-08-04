// vybe-test: java/method_calls/static_void_method_prints_uppercased_argument
// origin: languages/java/tests/java/test_method_calls.rs

public class Main {
static void announce(String msg) { System.out.println(msg.toUpperCase()); }
    public static void main(String[] args) {
announce("quiet");
    }
}

