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

// vybe-test: java/object_wait_notify/object_notify_before_wait_misses_signal
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class Miss {
            boolean signaled = false;
            synchronized void signal() { signaled = true; notify(); }
            synchronized void check() throws InterruptedException {
                if (!signaled) wait(50);
                System.out.println(signaled);
            }
        }
    public static void main(String[] args) throws Throwable {
Miss m = new Miss(); m.signal(); Thread t = new Thread(() -> { try { m.check(); } catch (InterruptedException e) {} }); t.start(); t.join();
    }
}

