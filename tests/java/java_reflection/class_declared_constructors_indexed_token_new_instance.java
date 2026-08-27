import java.lang.reflect.*;
public class Main {
    static class Foo { public Foo() {} }
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        Constructor<?> c = Foo.class.getDeclaredConstructors()[0];
        Object f = c.newInstance();
        __p(f != null);
        __check("true");
    }
}
