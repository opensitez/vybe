public class Main {
    static int echo(int x) { return x; }
    static int echo(int x, int y) { return x + y; }
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        __p(echo(1) + echo(2, 3));
        __check("6");
    }
}
