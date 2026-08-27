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

// vybe-test: java/object_wait_notify/object_wait_with_timeout_returns_when_notified
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class Timed {
            boolean done = false;
            synchronized void waitBriefly() throws InterruptedException {
                wait(500);
                if (done) System.out.println("notified");
            }
            synchronized void complete() { done = true; notify(); }
        }
    public static void main(String[] args) throws Throwable {
Timed t = new Timed(); Thread w = new Thread(() -> { try { t.waitBriefly(); } catch (InterruptedException e) {} }); w.start(); Thread.sleep(10); t.complete(); w.join();
    }
}

