// vybe-test: java/object_wait_notify/object_wait_notify_two_stage_pipeline
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class Stage {
            int stage = 0;
            synchronized void advance() throws InterruptedException { while (stage < 1) wait(); stage = 2; notify(); }
            synchronized void start() { stage = 1; notify(); }
            synchronized void finish() throws InterruptedException { while (stage < 2) wait(); System.out.println(stage); }
        }
    public static void main(String[] args) {
Stage s = new Stage(); Thread mid = new Thread(() -> { try { s.advance(); } catch (InterruptedException e) {} }); Thread end = new Thread(() -> { try { s.finish(); } catch (InterruptedException e) {} }); mid.start(); end.start(); Thread.sleep(10); s.start(); mid.join(); end.join();
    }
}

