// vybe-test: java/object_wait_notify/object_wait_releases_lock_for_other_thread
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class LockOrder {
            boolean flag = false;
            synchronized void setter() { flag = true; notify(); }
            synchronized void getter() throws InterruptedException {
                while (!flag) wait();
                System.out.println(flag);
            }
        }
    public static void main(String[] args) {
LockOrder lo = new LockOrder(); Thread t = new Thread(() -> { try { lo.getter(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); lo.setter(); t.join();
    }
}

