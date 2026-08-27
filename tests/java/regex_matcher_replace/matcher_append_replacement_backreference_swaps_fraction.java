import java.util.regex.*;
public class Main {
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        Pattern p = Pattern.compile("(\\d+)/(\\d+)");
        Matcher m = p.matcher("1/2");
        String res = m.replaceAll("$2/$1");
        __p(res);
        __check("2/1");
    }
}
