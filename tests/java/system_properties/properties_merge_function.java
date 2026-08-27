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
        Properties p = new Properties();
        p.put("k", "v1");
        p.merge("k", "v2", (v1, v2) -> String.valueOf(v1) + String.valueOf(v2));
        __p(p.get("k"));
        __check("v1v2");
    }
}
