// vybe-test: csharp/csharp_linq_let_join/let_clause_introduces_named_subexpression
// origin: languages/csharp/tests/csharp/test_csharp_linq_let_join.rs

using static __Harness;

var result =
    from s in new[]{"hello","hi","world"}
let len=s.Length
    where len>3
    select s;
foreach(var x in result) __P((x).ToString());
__Check("hello\nworld");

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
