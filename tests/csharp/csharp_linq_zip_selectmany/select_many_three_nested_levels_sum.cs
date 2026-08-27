// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_three_nested_levels_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

using static __Harness;

var data=new[]{new[]{new[]{1,2}},new[]{new[]{3}}}
;
var flat=data.SelectMany(a=>a).SelectMany(b=>b);
__P((flat.Sum()).ToString());
__Check("6");

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
