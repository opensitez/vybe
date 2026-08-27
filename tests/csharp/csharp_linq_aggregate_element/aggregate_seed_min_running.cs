// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_seed_min_running
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

using static __Harness;

var min=new[]{3,1,4,1,5}
.Aggregate(int.MaxValue,(acc,x)=>x<acc?x:acc);
__P((min).ToString());
__Check("1");

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
