// vybe-test: csharp/csharp_modern/object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

using static __Harness;

var p = new Point { X = 10, Y = 20 }
;
__P((p.X + p.Y).ToString());
__Check("30");

class Point {
    public int X { get; set; }
    public int Y { get; set; }
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
