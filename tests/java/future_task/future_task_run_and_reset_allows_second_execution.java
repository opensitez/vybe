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
        int[] count = new int[]{0};
        MyTask<Integer> task = new MyTask<>(() -> {
            count[0]++;
            return count[0];
        });
        task.run();
        task.runAndReset();
        task.run();
        __p(count[0]);
        __check("2");
    }
}
