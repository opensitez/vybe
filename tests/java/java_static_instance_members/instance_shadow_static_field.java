public class Main {
    static class Thing { static int x = 1; int instanceX = 2; }
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        Thing t = new Thing();
        __p(Thing.x + t.instanceX);
        __check("3");
    }
}
