// vybe-test: csharp/common_patterns/params_array_basic
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

Logger.Log("one", "two", "three");
__Check("one\ntwo\nthree");

class Logger {
    public static void Log(params string[] messages) {
        foreach (var m in messages) __P((m).ToString());
    }
}

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
