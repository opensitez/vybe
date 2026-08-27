// vybe-test: csharp/csharp_linq_complex/group_by_then_order_group_keys_ascending
// origin: languages/csharp/tests/csharp/test_csharp_linq_complex.rs

using static __Harness;

var items=new[]{(Cat:"b",Val:2),(Cat:"a",Val:1),(Cat:"b",Val:4),(Cat:"a",Val:3)}
;
var groups=items.GroupBy(i=>i.Cat).OrderBy(g=>g.Key)
    .Select(g=>(g.Key,g.Sum(i=>i.Val)));
foreach(var(k,s) in groups) __P(($"{k}:{s}").ToString());
__Check("a:4\nb:6");

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
