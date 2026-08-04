// vybe-test: java/abstract_classes/abstract_void_method_side_effect
// origin: languages/java/tests/java/test_abstract_classes.rs

public class Main {
static abstract class Logger { abstract void log(String msg); }
        static class PrintLogger extends Logger { void log(String msg) { System.out.println(msg); } }
    public static void main(String[] args) {
Logger l = new PrintLogger(); l.log("trace");
    }
}

