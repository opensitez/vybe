// vybe-test: csharp/csharp_linq_skip_take_distinct/distinct_by_record_key_first_values
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

using static __Harness;

var r=new[]{(K:1,V:"a"),(K:1,V:"b"),(K:2,V:"c")}
.DistinctBy(t=>t.K);
__P((r.First().V).ToString());
__P((r.Last().V).ToString());
__Check("a\nc");

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
