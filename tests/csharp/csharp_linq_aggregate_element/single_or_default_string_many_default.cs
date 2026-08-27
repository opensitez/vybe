// vybe-test: csharp/csharp_linq_aggregate_element/single_or_default_string_many_default
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

using static __Harness;

try { __P((new[]{"a","b"}.SingleOrDefault("z")).ToString()); }
catch (InvalidOperationException) { __P("threw"); }
__Check("threw");

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
