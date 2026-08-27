// vybe-test: csharp/csharp_linq_groupby_join/join_links_two_sequences_on_matching_keys
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_join.rs

using static __Harness;

var ids  = new[] { 1, 2, 3 }
;
var names = new[] { (Id:1, Name:"one"), (Id:2, Name:"two") }
;
var joined = ids.Join(names, id => id, n => n.Id, (id, n) => n.Name);
foreach (var s in joined) __P((s).ToString());
__Check("one\ntwo");

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
