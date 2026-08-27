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

// vybe-test: java/object_wait_notify/object_wait_notify_integer_state_machine
// origin: languages/java/tests/java/test_object_wait_notify.rs

public class Main {
static class FSM {
            int state = 0;
            synchronized void toState(int s) throws InterruptedException { while (state < s - 1) wait(); state = s; notifyAll(); }
            synchronized void report(int s) throws InterruptedException { while (state < s) wait(); System.out.println(state); }
        }
    public static void main(String[] args) throws Throwable {
FSM f = new FSM(); Thread t = new Thread(() -> { try { f.report(2); } catch (InterruptedException e) {} }); t.start(); Thread.sleep(10); try { f.toState(1); f.toState(2); } catch (InterruptedException e) {} t.join();
    }
}

