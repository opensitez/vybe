// vybe-test: csharp/csharp_struct_features/readonly_struct_field_cannot_be_mutated_but_is_readable
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

using static __Harness;

var obj = new Immutable(7);
__P((obj.Value).ToString());
__Check("7");

readonly struct Immutable { public readonly int Value; public Immutable(int v) { Value=v; } }

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
