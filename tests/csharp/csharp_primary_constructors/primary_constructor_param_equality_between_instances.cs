// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_equality_between_instances
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

var a = new Id(5);
var b = new Id(5);
__P((a == b).ToString());
__Check("True");

record Id(int Value);

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
