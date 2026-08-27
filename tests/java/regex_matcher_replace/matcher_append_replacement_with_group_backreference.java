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
        Pattern p = Pattern.compile("([a-z]+)");
        Matcher m = p.matcher("abc");
        String res = m.replaceAll("[$1]");
        __p(res);
        __check("[abc]");
    }
}
