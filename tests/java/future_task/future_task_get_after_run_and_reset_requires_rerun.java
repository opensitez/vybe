import java.util.concurrent.*;
public class Main {
    static class MyTask<V> extends FutureTask<V> {
        public MyTask(Callable<V> c) { super(c); }
        public boolean runAndReset() { return super.runAndReset(); }
    }
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        MyTask<Integer> task = new MyTask<>(() -> 42);
        task.run();
        task.runAndReset();
        __p(task.isDone());
        __check("false");
    }
}
