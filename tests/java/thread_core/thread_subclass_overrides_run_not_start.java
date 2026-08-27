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

// vybe-test: java/thread_core/thread_subclass_overrides_run_not_start
// origin: languages/java/tests/java/test_thread_core.rs

public class Main {
static class Safe extends Thread {
            public void run() { System.out.println("body"); }
        }
    public static void main(String[] args) throws Throwable {
Safe s = new Safe(); s.start(); s.join();
    }
}

