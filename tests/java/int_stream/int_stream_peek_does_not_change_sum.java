import java.util.concurrent.atomic.*;
import java.util.stream.*;
public class Main {
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        AtomicInteger count = new AtomicInteger();
        int sum = IntStream.of(1, 2, 3).peek(x -> count.incrementAndGet()).sum();
        __p(count.get() + sum);
        __check("9");
    }
}
