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
        TreeSet<Integer> set = new TreeSet<>(Arrays.asList(1, 2, 3, 4, 5));
        NavigableSet<Integer> view = set.tailSet(2, true);
        __p(view.floor(3));
        __check("3");
    }
}
