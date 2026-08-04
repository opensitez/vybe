// vybe-test: java/object_wait_notify/object_wait_zero_timeout_returns_after_notify
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class Zero {
            boolean set = false;
            synchronized void waitZero() throws InterruptedException { while (!set) wait(0); System.out.println("ok"); }
            synchronized void mark() { set = true; notify(); }
        }
    public static void main(String[] args) {
Zero z = new Zero(); Thread t = new Thread(() -> { try { z.waitZero(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); z.mark(); t.join();
    }
}

