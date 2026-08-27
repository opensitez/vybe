public class Main {
    static class ArrBox {
        int[] arr;
        ArrBox(int[] arr) { this.arr = arr; }
        int getTotal() { return arr[0]; }
    }
    static String __buf = "";
    static void __p(Object o) { __buf = __buf + String.valueOf(o) + "\n"; }
    static void __check(String want) {
        String got = __buf;
        if (got.endsWith("\n")) got = got.substring(0, got.length() - 1);
        if (!got.equals(want)) throw new RuntimeException("fail: " + got);
    }
    public static void main(String[] args) throws Throwable {
        ArrBox b = new ArrBox(new int[]{5});
        __p(b.getTotal());
        __check("5");
    }
}
