// vybe-test: csharp/csharp_oop/interface_with_multiple_methods
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

using static __Harness;

var c = new Calc();
__P((c.Add(3, 4)).ToString());
__P((c.Multiply(3, 4)).ToString());
__Check("7\n12");

interface ICalculator {
    int Add(int a, int b);
    int Multiply(int a, int b);
}

class Calc : ICalculator {
    public int Add(int a, int b) { return a + b; }
    public int Multiply(int a, int b) { return a * b; }
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
