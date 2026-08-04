// vybe-test: java/object_wait_notify/object_wait_spurious_wakeup_guarded_by_while
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class Guard {
            int ready = 0;
            synchronized void await() throws InterruptedException { while (ready == 0) wait(); System.out.println(ready); }
            synchronized void setReady(int v) { ready = v; notifyAll(); }
        }
    public static void main(String[] args) {
Guard g = new Guard(); Thread t = new Thread(() -> { try { g.await(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); g.setReady(5); t.join();
    }
}

