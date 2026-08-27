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

// vybe-test: java/object_wait_notify/object_wait_notify_all_resets_condition
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class Reset {
            boolean active = true;
            synchronized void deactivate() { active = false; notifyAll(); }
            synchronized void waitInactive() throws InterruptedException { while (active) wait(); System.out.println("inactive"); }
        }
    public static void main(String[] args) throws Throwable {
Reset r = new Reset(); Thread t = new Thread(() -> { try { r.waitInactive(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); r.deactivate(); t.join();
    }
}

