// vybe-test: csharp/csharp_linq_projections/to_dictionary_builds_map_from_sequence
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

using static __Harness;

var dict = new[]{"a","bb","ccc"}
.ToDictionary(s => s, s => s.Length);
__P((dict["bb"]).ToString());
__Check("2");

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
