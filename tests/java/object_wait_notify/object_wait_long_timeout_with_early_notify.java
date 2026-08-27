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

// vybe-test: java/object_wait_notify/object_wait_long_timeout_with_early_notify
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class Early {
            boolean hit = false;
            synchronized void waitLong() throws InterruptedException { while (!hit) wait(1000); System.out.println("early"); }
            synchronized void ping() { hit = true; notify(); }
        }
    public static void main(String[] args) throws Throwable {
Early e = new Early(); Thread t = new Thread(() -> { try { e.waitLong(); } catch (InterruptedException e2) {} }); t.start(); Thread.sleep(10); e.ping(); t.join();
    }
}

