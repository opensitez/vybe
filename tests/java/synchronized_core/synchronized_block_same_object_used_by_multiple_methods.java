// vybe-test: java/synchronized_core/synchronized_block_same_object_used_by_multiple_methods
// origin: languages/java/tests/java/test_synchronized_core.rs

public class Main {

    // A static String, NOT a StringBuilder. Calling a method on a bare static
    // FIELD receiver fails under Vybe with "undefined is not callable"
    // (measured): `SB.append(x)` throws while `StringBuilder l = SB;
    // l.append(x)` works, so the method is resolved from the receiver's
    // declared type at the call site and a static field carries none. String
    // concatenation onto a static field has no such problem.
    static String __buf = "";

    static void __p(Object o) {
        __buf = __buf + String.valueOf(o) + "\n";
    }

    static void __pr(Object o) {
        __buf = __buf + String.valueOf(o);
    }

    static void __check(String want) {
        String got = __buf;
        // The final `println` contributes a trailing newline that the expected
        // line vector never carried, so it is not part of the comparison.
        if (got.endsWith("\n")) {
            got = got.substring(0, got.length() - 1);
        }
        if (!got.equals(want)) {
            System.out.println("FAIL: want [" + want + "] got [" + got + "]");
            throw new RuntimeException("assertion failed");
        }
    }

static class Account {
            final Object lock = new Object();
            int balance = 0;
            void deposit(int amount) {
                synchronized (lock) { balance += amount; }
            }
            int snapshot() {
                synchronized (lock) { return balance; }
            }
        }
    public static void main(String[] args) {
Account a = new Account(); Thread t1 = new Thread(() -> a.deposit(5)); Thread t2 = new Thread(() -> a.deposit(7)); t1.start(); t2.start(); t1.join(); t2.join(); __p(a.snapshot());
__check("12");
    }
}

