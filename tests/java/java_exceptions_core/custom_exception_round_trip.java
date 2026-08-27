public class Main {
    static class MyEx extends Exception { MyEx(String s) { super(s); } }
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        try { throw new MyEx("err"); } catch (MyEx e) { __p(e.getMessage()); }
        __check("err");
    }
}
