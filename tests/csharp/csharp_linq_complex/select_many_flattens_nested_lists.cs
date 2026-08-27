// vybe-test: csharp/csharp_linq_complex/select_many_flattens_nested_lists
// origin: languages/csharp/tests/csharp/test_csharp_linq_complex.rs

using static __Harness;

var data=new[]{
    new[]{1,2,3},
    new[]{4,5},
    new[]{6}
}
;
int sum=data.SelectMany(x=>x).Sum();
__P((sum).ToString());
__Check("21");

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
