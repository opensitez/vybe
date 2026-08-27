import java.util.*;
import java.util.stream.*;
import java.util.function.*;
import java.util.concurrent.*;
import java.time.*;
import java.time.format.*;
import java.net.*;
import java.io.*;
import java.nio.file.*;
import java.lang.reflect.*;

// vybe-test: java/object_wait_notify/object_notify_only_one_waiter_awakened
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class Counter {
            int n = 0;
            synchronized void inc() throws InterruptedException {
                while (n == 0) wait();
                n--;
                System.out.println("go");
            }
            synchronized void signal() { n = 1; notify(); }
        }
    public static void main(String[] args) throws Throwable {
Counter c = new Counter(); Thread t = new Thread(() -> { try { c.inc(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); c.signal(); t.join();
    }
}

