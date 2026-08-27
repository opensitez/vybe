// vybe-test: csharp/csharp_static_classes/static_method_can_call_other_static_methods_in_same_class
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

using static __Harness;

__P((Calc.Sum3(1,2,3)).ToString());
__Check("6");

static class Calc {
    static int Add(int a, int b) => a+b;
    public static int Sum3(int a, int b, int c) => Add(Add(a,b),c);
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
