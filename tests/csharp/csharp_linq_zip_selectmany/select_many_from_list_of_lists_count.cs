// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_from_list_of_lists_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

using static __Harness;

var lists=new System.Collections.Generic.List<int[]>{
    new[]{1,2},new[]{3}}
;
__P((lists.SelectMany(x=>x).Count()).ToString());
__Check("3");

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
