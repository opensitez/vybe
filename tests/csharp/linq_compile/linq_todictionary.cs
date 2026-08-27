// vybe-test: csharp/linq_compile/linq_todictionary
// origin: languages/csharp/tests/csharp/test_linq_compile.rs
// vybe-test-mode: compile

using static __Harness;

var words = new List<string>();
words.Add("a");
words.Add("bb");
var dict = words.ToDictionary(w => w);

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
