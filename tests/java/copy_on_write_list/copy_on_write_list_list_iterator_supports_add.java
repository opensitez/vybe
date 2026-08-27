import java.util.concurrent.*;
import java.util.*;
public class Main {
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        CopyOnWriteArrayList<String> list = new CopyOnWriteArrayList<>(Arrays.asList("a"));
        ListIterator<String> it = list.listIterator();
        try {
            it.add("x");
        } catch (UnsupportedOperationException e) {
            __p("uoe");
        }
        __check("uoe");
    }
}
