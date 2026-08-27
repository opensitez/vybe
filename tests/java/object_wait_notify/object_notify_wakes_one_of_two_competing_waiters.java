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

// vybe-test: java/object_wait_notify/object_notify_wakes_one_of_two_competing_waiters
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class Compete {
            int tickets = 0;
            synchronized void take() throws InterruptedException { while (tickets == 0) wait(); tickets--; System.out.println("got"); }
            synchronized void give() { tickets++; notify(); }
        }
    public static void main(String[] args) throws Throwable {
Compete c = new Compete(); Thread t = new Thread(() -> { try { c.take(); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); c.give(); t.join();
    }
}

