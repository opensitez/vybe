// vybe-test: csharp/modern_features/null_conditional_chain
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

var o = new Outer();
__P((o.Child?.Value ?? "missing").ToString());
o.Child = new Inner();
__P((o.Child?.Value ?? "missing").ToString());
__Check("missing\nfound");

class Inner { public string Value = "found"; }

class Outer { public Inner Child; }

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
