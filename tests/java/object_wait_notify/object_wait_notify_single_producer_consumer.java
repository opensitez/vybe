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

// vybe-test: java/object_wait_notify/object_wait_notify_single_producer_consumer
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class Box {
            int value = 0;
            synchronized void produce() { value = 42; notify(); }
            synchronized void consume() throws InterruptedException {
                while (value == 0) wait();
                System.out.println(value);
            }
        }
    public static void main(String[] args) throws Throwable {
Box box = new Box(); Thread consumer = new Thread(() -> { try { box.consume(); } catch (InterruptedException e) {} }); consumer.start(); Thread.sleep(10); box.produce(); consumer.join();
    }
}

