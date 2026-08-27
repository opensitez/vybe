public class Main {
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    static <A, B> String join(A a, B b) { return String.valueOf(a) + String.valueOf(b); }
    public static void main(String[] args) throws Throwable {
        __p(join("a", 1));
        __check("a1");
    }
}
