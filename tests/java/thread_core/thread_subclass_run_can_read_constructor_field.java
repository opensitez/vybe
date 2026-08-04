// vybe-test: java/thread_core/thread_subclass_run_can_read_constructor_field
// origin: languages/java/tests/java/test_thread_core.rs

public class Main {
static class EchoThread extends Thread {
            String msg;
            EchoThread(String msg) { this.msg = msg; }
            public void run() { System.out.println(msg); }
        }
    public static void main(String[] args) {
EchoThread t = new EchoThread("payload"); t.start(); t.join();
    }
}

