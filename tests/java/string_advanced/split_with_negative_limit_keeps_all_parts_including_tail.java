public class Main {
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        String[] parts = "a,b,".split(",", -1);
        __p(parts.length);
        __p(parts[0]);
        __p(parts[1]);
        __p(parts[2]);
        __check("3\na\nb\n");
    }
}
