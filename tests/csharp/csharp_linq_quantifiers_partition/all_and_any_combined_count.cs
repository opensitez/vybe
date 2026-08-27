// vybe-test: csharp/csharp_linq_quantifiers_partition/all_and_any_combined_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

using static __Harness;

var data=new[]{2,4,6,8}
;
__P((data.All(x=>x%2==0)?1:0).ToString());
__P((data.Any(x=>x>5)?1:0).ToString());
__Check("1\n1");

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
