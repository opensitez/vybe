// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_repeat_each_element_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

using static __Harness;

var flat=new[]{1,2}
.SelectMany(n=>new[]{n,n,n});
__P((flat.Sum()).ToString());
__Check("9");

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
