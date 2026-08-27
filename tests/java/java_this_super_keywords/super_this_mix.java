
public class Main {
    static class Base { int x = 1; }
    static class Sub extends Base { int x = 2; int getSuper() { return super.x; } int getThis() { return this.x; } }


    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __pr(Object o) { __buf = __buf + String.valueOf(o); }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) {
            System.out.println("FAIL: want [" + want + "] got [" + got + "]");
            throw new RuntimeException("assertion failed");
        }
    }
    public static void main(String[] args) throws Throwable {
        Sub s = new Sub();
        __p(s.getSuper() + s.getThis());
        __check("3");
    }
}
