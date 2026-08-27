public class Main {
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        byte[] b = "hi".getBytes(java.nio.charset.StandardCharsets.UTF_8);
        __p(b[0]);
        __p(b[1]);
        __check("104\n105");
    }
}
