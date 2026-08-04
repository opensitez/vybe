// vybe-test: java/generic_core/generic_void_method_prints_argument
// origin: languages/java/tests/java/test_generic_core.rs

public class Main {
static class Echo {
            static <T> void show(T value) { System.out.println(value); }
        }
    public static void main(String[] args) {
Echo.show("ping");
    }
}

