// vybe-test: java/thread_core/thread_subclass_overrides_run_not_start
// origin: languages/java/tests/java/test_thread_core.rs

public class Main {
static class Safe extends Thread {
            public void run() { System.out.println("body"); }
        }
    public static void main(String[] args) {
Safe s = new Safe(); s.start(); s.join();
    }
}

