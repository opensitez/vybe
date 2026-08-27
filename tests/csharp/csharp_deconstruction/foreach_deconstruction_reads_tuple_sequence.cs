// vybe-test: csharp/csharp_deconstruction/foreach_deconstruction_reads_tuple_sequence
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

using static __Harness;

var pairs = new[] { ("a", 1), ("b", 2) }
;
foreach (var (letter, number) in pairs) {
    __P((letter + number).ToString());
}
__Check("a1\nb2");

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
