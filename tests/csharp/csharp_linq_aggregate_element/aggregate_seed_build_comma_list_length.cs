// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_seed_build_comma_list_length
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

using static __Harness;

var text=new[]{1,2,3}
.Aggregate("",(acc,x)=>acc==""?x.ToString():acc+","+x);
__P((text.Length).ToString());
__Check("5");

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
