import java.util.*;
public class Main {
    static class C<T extends Comparable<T>> implements Comparator<T> {
        public int compare(T a, T b) { return a.compareTo(b); }
    }
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        C<Integer> c = new C<>();
        __p(c.compare(1, 2) < 0);
        __check("true");
    }
}
