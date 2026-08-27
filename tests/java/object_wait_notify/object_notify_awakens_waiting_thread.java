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

// vybe-test: java/object_wait_notify/object_notify_awakens_waiting_thread
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class Gate {
            boolean open = false;
            synchronized void awaitOpen() throws InterruptedException {
                while (!open) wait();
                System.out.println("open");
            }
            synchronized void open() { open = true; notify(); }
        }
    public static void main(String[] args) throws Throwable {
Gate g = new Gate(); Thread t = new Thread(() -> { try { g.awaitOpen(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); g.open(); t.join();
    }
}

