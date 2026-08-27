// vybe-test: csharp/csharp_linq_advanced/max_by_returns_element_with_maximum_key
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

using static __Harness;

var words=new[]{"a","bbb","cc"}
;
__P((words.MaxBy(w=>w.Length)).ToString());
__Check("bbb");

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
