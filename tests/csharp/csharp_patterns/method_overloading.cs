// vybe-test: csharp/csharp_patterns/method_overloading
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

using static __Harness;

var p = new Printer();
__P((p.Print(42)).ToString());
__P((p.Print("hi")).ToString());
__P((p.Print(1, 2)).ToString());
__Check("int:42\nstr:hi\npair:1,2");

class Printer {
    public string Print(int x) { return "int:" + x; }
    public string Print(string x) { return "str:" + x; }
    public string Print(int x, int y) { return "pair:" + x + "," + y; }
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
