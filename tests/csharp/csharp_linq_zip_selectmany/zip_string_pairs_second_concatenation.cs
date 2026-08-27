// vybe-test: csharp/csharp_linq_zip_selectmany/zip_string_pairs_second_concatenation
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

using static __Harness;

var z=new[]{"a","b"}
.Zip(new[]{"1","2"},(x,y)=>x+y);
__P((z.Last()).ToString());
__Check("b2");

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
