// vybe-test: csharp/csharp_linq_numeric/max_by_returns_whole_element_not_just_key
// origin: languages/csharp/tests/csharp/test_csharp_linq_numeric.rs

using static __Harness;

var words=new[]{"cat","elephant","ox"}
;
__P((words.MaxBy(w=>w.Length)).ToString());
__Check("elephant");

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
