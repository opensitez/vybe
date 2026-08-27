import java.util.*;
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
        List<String> list = Arrays.asList("a", "bb", "ccc");
        int total = list.stream().collect(Collectors.reducing(0, str -> str.length(), Integer::sum));
        __p(total);
        __check("6");
    }
}
