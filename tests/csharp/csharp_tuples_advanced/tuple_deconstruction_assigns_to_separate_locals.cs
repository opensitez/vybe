// vybe-test: csharp/csharp_tuples_advanced/tuple_deconstruction_assigns_to_separate_locals
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

using static __Harness;

var (a, b, c) = (10, 20, 30);
__P((a+b+c).ToString());
__Check("60");

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
