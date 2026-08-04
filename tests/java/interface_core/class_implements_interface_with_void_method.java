// vybe-test: java/interface_core/class_implements_interface_with_void_method
// origin: languages/java/tests/java/test_interface_core.rs

public class Main {
interface Sink { void accept(int n); }
        static class PrintSink implements Sink {
            public void accept(int n) { System.out.println(n); }
        }
    public static void main(String[] args) {
Sink s = new PrintSink(); s.accept(42);
    }
}

