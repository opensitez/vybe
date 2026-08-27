// vybe-test: csharp/csharp_oop/sealed_class
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

using static __Harness;

var s = new Singleton();
__P((s.Value).ToString());
__Check("42");

sealed class Singleton {
    public int Value = 42;
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
