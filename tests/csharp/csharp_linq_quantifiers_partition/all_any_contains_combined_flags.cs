// vybe-test: csharp/csharp_linq_quantifiers_partition/all_any_contains_combined_flags
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

using static __Harness;

var xs=new[]{1,2,3}
;
__P((xs.All(x=>x>0)?1:0).ToString());
__P((xs.Any(x=>x==2)?1:0).ToString());
__P((xs.Contains(4)?1:0).ToString());
__Check("1\n1\n0");

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
