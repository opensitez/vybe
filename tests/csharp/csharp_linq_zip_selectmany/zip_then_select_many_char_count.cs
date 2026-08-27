// vybe-test: csharp/csharp_linq_zip_selectmany/zip_then_select_many_char_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

using static __Harness;

var words=new[]{"hi","go"}
;
var letters=words.Zip(new[]{1,2},(w,n)=>w).SelectMany(w=>w);
__P((letters.Count()).ToString());
__Check("4");

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
