// vybe-test: java/object_wait_notify/object_wait_notify_counter_increment_while_waiting
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class Acc {
            int total = 0;
            synchronized void add(int n) { total += n; if (total >= 10) notify(); }
            synchronized void waitTen() throws InterruptedException { while (total < 10) wait(); System.out.println(total); }
        }
    public static void main(String[] args) {
Acc a = new Acc(); Thread t = new Thread(() -> { try { a.waitTen(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); for (int i = 0; i < 5; i++) a.add(2); t.join();
    }
}

