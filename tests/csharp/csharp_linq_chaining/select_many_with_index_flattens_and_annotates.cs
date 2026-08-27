// vybe-test: csharp/csharp_linq_chaining/select_many_with_index_flattens_and_annotates
// origin: languages/csharp/tests/csharp/test_csharp_linq_chaining.rs

using static __Harness;

var groups=new[]{new[]{1,2},new[]{3,4}}
;
var result=groups.SelectMany((g,i)=>g.Select(x=>i*10+x));
__P((string.Join(",",result)).ToString());
__Check("1,2,13,14");

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
