// vybe-test: csharp/csharp_linq_zip_selectmany/zip_three_way_manual_via_select_index
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

using static __Harness;

var a=new[]{1,2,3}
;
var b=new[]{4,5,6}
;
var z=a.Zip(b,(x,y)=>x+y);
__P((z.ElementAt(1)).ToString());
__Check("7");

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
