// vybe-test: csharp/more_classes/record_tostring
// origin: languages/csharp/tests/csharp/test_more_classes.rs

using static __Harness;

var p = new Point(3, 7);
__P((p.Display()).ToString());
__Check("Point(3, 7)");

record Point(int X, int Y) {
            public string Display() {
                return "Point(" + X + ", " + Y + ")";
            }
        }

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
