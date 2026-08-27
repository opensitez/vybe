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
        List<Integer> list = Arrays.asList(3, 1, 2);
        Collections.sort(list, Comparator.naturalOrder());
        __p(list.get(0));
        __check("1");
    }
}
