// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_strings_to_chars_first
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

using static __Harness;

var chars=new[]{"ab","c"}
.SelectMany(s=>s);
__P((chars.First()).ToString());
__Check("a");

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
