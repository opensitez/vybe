// vybe-test: csharp/csharp_deconstruction/tuple_deconstruction_assigns_two_scalars
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

using static __Harness;

var (x, y) = (3, 4);
__P((x).ToString());
__P((y).ToString());
__Check("3\n4");

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
