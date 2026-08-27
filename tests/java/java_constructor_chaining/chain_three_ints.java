public class Main {
    static class Triple {
        int a, b, c;
        Triple(int a) { this(a, 0, 0); }
        Triple(int a, int b, int c) { this.a = a; this.b = b; this.c = c; }
    }
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        Triple t = new Triple(1);
        __p(t.a);
        __check("1");
    }
}
