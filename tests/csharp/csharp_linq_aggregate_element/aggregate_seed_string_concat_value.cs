// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_seed_string_concat_value
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

using static __Harness;

var s=new[]{"a","b","c"}
.Aggregate("",(acc,x)=>acc+x);
__P((s).ToString());
__Check("abc");

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
