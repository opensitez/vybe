// vybe-test: csharp/csharp_arrays_advanced/array_of_strings
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

using static __Harness;

var words = new[] { "hello", "world" }
;
__P((words[0] + " " + words[1]).ToString());
__P((words.Length).ToString());
__Check("hello world\n2");

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
