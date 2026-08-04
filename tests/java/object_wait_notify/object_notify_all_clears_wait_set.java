// vybe-test: java/object_wait_notify/object_notify_all_clears_wait_set
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class Broadcast {
            int gen = 0;
            synchronized void waitGen(int g) throws InterruptedException { while (gen < g) wait(); System.out.println(gen); }
            synchronized void bump() { gen++; notifyAll(); }
        }
    public static void main(String[] args) {
Broadcast b = new Broadcast(); Thread t = new Thread(() -> { try { b.waitGen(1); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); b.bump(); t.join();
    }
}

