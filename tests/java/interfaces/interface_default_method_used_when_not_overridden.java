// vybe-test: java/interfaces/interface_default_method_used_when_not_overridden
// origin: languages/java/tests/java/test_interfaces.rs

public class Main {
interface Logger { default void log(String msg) { System.out.println(msg); } }
        static class ConsoleLogger implements Logger {}
    public static void main(String[] args) {
Logger l = new ConsoleLogger(); l.log("ok");
    }
}

