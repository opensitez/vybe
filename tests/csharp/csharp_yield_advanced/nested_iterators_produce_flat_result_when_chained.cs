// vybe-test: csharp/csharp_yield_advanced/nested_iterators_produce_flat_result_when_chained
// origin: languages/csharp/tests/csharp/test_csharp_yield_advanced.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Doubles(int n){
    yield return n; yield return n*2;
}
var result=new[]{1,2,3}
.SelectMany(Doubles);
__P((string.Join(",",result)).ToString());
__Check("1,2,2,4,3,6");

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
