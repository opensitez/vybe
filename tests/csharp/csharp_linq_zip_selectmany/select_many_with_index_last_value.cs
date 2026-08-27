// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_with_index_last_value
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

using static __Harness;

var flat=new[]{new[]{10},new[]{20,30}}
.SelectMany((x,i)=>x.Select(v=>v+i));
__P((flat.Last()).ToString());
__Check("31");

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
