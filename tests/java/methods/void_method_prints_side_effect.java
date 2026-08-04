// vybe-test: java/methods/void_method_prints_side_effect
// origin: languages/java/tests/java/test_methods.rs

public class Main {
static void shout(String msg) { System.out.println(msg.toUpperCase()); }
    public static void main(String[] args) {
shout("quiet");
    }
}

