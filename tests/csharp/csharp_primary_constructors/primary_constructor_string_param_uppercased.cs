// vybe-test: csharp/csharp_primary_constructors/primary_constructor_string_param_uppercased
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

__P((new Shout("go").Loud()).ToString());
__Check("GO");

class Shout(string word) { public string Loud() => word.ToUpper(); }

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
