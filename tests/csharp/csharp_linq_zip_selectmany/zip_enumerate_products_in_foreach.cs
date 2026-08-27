// vybe-test: csharp/csharp_linq_zip_selectmany/zip_enumerate_products_in_foreach
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

using static __Harness;

var z=new[]{2,3}
.Zip(new[]{4,5},(a,b)=>a*b);
foreach(var n in z) __P((n).ToString());
__Check("8\n15");

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
