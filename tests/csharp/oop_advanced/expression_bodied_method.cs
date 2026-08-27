// vybe-test: csharp/oop_advanced/expression_bodied_method
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

var c = new Calc();
__P((c.Square(7)).ToString());
__P((c.Greet("World")).ToString());
__Check("49\nHello World");

class Calc {
    public int Square(int x) => x * x;
    public string Greet(string name) => "Hello " + name;
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
