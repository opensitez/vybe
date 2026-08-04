// vybe-test: java/object_wait_notify/object_notify_without_waiter_is_noop
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class NoWaiter {
            synchronized void ping() { notify(); System.out.println("ping"); }
        }
    public static void main(String[] args) {
NoWaiter n = new NoWaiter(); n.ping();
    }
}

