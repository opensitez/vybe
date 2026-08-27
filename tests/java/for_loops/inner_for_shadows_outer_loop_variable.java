public class Main {
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        int sum = 0;
        for (int i = 0; i < 2; i++) {
            for (int k = 0; k < 2; k++) {
                sum += k;
            }
        }
        __p(sum);
        __check("2");
    }
}
