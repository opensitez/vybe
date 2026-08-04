// vybe-test: java/interface_core/unresolved_default_method_diamond_is_rejected
// origin: languages/java/tests/java/test_interface_core.rs
// vybe-test-mode: compile

public class Main {
    public static void main(String[] args) {
interface A { default String who() { return "a"; } } interface B { default String who() { return "b"; } } class C implements A, B { } public class Main { public static void main(String[] a) { System.out.println(new C().who()); } }
    }
}

