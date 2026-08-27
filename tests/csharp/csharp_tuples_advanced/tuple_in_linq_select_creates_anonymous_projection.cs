// vybe-test: csharp/csharp_tuples_advanced/tuple_in_linq_select_creates_anonymous_projection
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

using static __Harness;

var items = new[]{"apple","kiwi","pear"}
;
var proj = items.Select(s => (Name: s, Len: s.Length));
foreach(var x in proj) __P((x.Len).ToString());
__Check("5\n4\n4");

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
