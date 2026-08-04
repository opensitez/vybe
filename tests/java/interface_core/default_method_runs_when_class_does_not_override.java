// vybe-test: java/interface_core/default_method_runs_when_class_does_not_override
// origin: languages/java/tests/java/test_interface_core.rs

public class Main {
interface Logger { default void log(String msg) { System.out.println(msg); } }
        static class ConsoleLogger implements Logger {}
    public static void main(String[] args) {
Logger l = new ConsoleLogger(); l.log("ok");
    }
}

