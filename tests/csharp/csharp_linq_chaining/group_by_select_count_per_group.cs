// vybe-test: csharp/csharp_linq_chaining/group_by_select_count_per_group
// origin: languages/csharp/tests/csharp/test_csharp_linq_chaining.rs

using static __Harness;

var words=new[]{"cat","car","bar","bat","can"}
;
var groups=words.GroupBy(w=>w[0])
    .Select(g=>(g.Key,g.Count()))
    .OrderBy(t=>t.Key);
foreach(var(k,c) in groups) __P(($"{k}:{c}").ToString());
__Check("b:2\nc:3");

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
