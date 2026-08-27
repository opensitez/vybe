// vybe-test: csharp/common_patterns/as_operator
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

object obj = "hello";
string s = obj as string;
__P((s != null ? s : "null").ToString());
int? i = obj as int?;
__P((i != null ? i.ToString() : "null").ToString());
__Check("hello\nnull");

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
