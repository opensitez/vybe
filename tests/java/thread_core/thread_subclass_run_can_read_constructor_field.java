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

// vybe-test: java/thread_core/thread_subclass_run_can_read_constructor_field
// origin: languages/java/tests/java/test_thread_core.rs

public class Main {
static class EchoThread extends Thread {
            String msg;
            EchoThread(String msg) { this.msg = msg; }
            public void run() { System.out.println(msg); }
        }
    public static void main(String[] args) throws Throwable {
EchoThread t = new EchoThread("payload"); t.start(); t.join();
    }
}

