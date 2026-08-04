// vybe-test: java/thread_core/thread_subclass_run_invoked_directly_without_start
// origin: languages/java/tests/java/test_thread_core.rs

public class Main {
static class Worker extends Thread {
            public void run() { System.out.println("direct"); }
        }
    public static void main(String[] args) {
Worker w = new Worker(); w.run();
    }
}

