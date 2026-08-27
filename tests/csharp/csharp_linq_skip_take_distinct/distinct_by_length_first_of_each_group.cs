// vybe-test: csharp/csharp_linq_skip_take_distinct/distinct_by_length_first_of_each_group
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

using static __Harness;

var r=new[]{"a","bb","c","dd","eee"}
.DistinctBy(s=>s.Length);
foreach(var s in r) __P((s).ToString());
__Check("a\nbb\neee");

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
