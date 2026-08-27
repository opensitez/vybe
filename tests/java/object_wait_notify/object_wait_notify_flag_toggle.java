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

// vybe-test: java/object_wait_notify/object_wait_notify_flag_toggle
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class Toggle {
            boolean on = false;
            synchronized void turnOn() { on = true; notifyAll(); }
            synchronized void waitOn() throws InterruptedException { while (!on) wait(); System.out.println(on); }
        }
    public static void main(String[] args) throws Throwable {
Toggle t = new Toggle(); Thread w = new Thread(() -> { try { t.waitOn(); } catch (InterruptedException e) {} }); w.start(); Thread.sleep(10); t.turnOn(); w.join();
    }
}

