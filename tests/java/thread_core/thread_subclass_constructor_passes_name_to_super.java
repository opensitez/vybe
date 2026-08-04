// vybe-test: java/thread_core/thread_subclass_constructor_passes_name_to_super
// origin: languages/java/tests/java/test_thread_core.rs

public class Main {
static class NamedWorker extends Thread {
            NamedWorker(String name) { super(name); }
            public void run() { System.out.println(getName()); }
        }
    public static void main(String[] args) {
NamedWorker w = new NamedWorker("named"); w.start(); w.join();
    }
}

