import java.util.*;
import java.util.function.*;
public class Main {
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        List<Supplier<Integer>> tasks = new ArrayList<>();
        for (int i = 0; i < 3; i++) {
            final int fi = i;
            tasks.add(() -> fi);
        }
        __p(tasks.get(0).get() + tasks.get(2).get());
        __check("2");
    }
}
