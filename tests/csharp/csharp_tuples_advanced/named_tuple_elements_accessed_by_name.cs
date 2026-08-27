// vybe-test: csharp/csharp_tuples_advanced/named_tuple_elements_accessed_by_name
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

using static __Harness;

var p = (X: 3, Y: 4);
__P((p.X).ToString());
__P((p.Y).ToString());
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
