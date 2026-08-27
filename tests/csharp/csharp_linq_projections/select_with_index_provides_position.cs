// vybe-test: csharp/csharp_linq_projections/select_with_index_provides_position
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

using static __Harness;

var result = new[]{"a","b","c"}
.Select((x,i) => $"{i}:{x}");
foreach(var s in result) __P((s).ToString());
__Check("0:a\n1:b\n2:c");

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
