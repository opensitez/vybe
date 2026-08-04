// vybe-test: java/object_wait_notify/object_wait_on_private_lock_object
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class PrivateLock {
            final Object lock = new Object();
            boolean ready = false;
            void await() throws InterruptedException {
                synchronized (lock) { while (!ready) lock.wait(); System.out.println("ready"); }
            }
            void setReady() { synchronized (lock) { ready = true; lock.notify(); } }
        }
    public static void main(String[] args) {
PrivateLock p = new PrivateLock(); Thread t = new Thread(() -> { try { p.await(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); p.setReady(); t.join();
    }
}

