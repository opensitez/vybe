// vybe-test: csharp/csharp_linq_zip_selectmany/zip_result_selector_returns_tuple_like_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

using static __Harness;

var z=new[]{1,2}
.Zip(new[]{3,4},(a,b)=>a*10+b);
__P((z.Sum()).ToString());
__Check("37");

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
