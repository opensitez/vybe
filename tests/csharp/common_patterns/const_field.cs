// vybe-test: csharp/common_patterns/const_field
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

__P((Config.MaxRetries).ToString());
__P((Config.AppName).ToString());
__Check("3\nMyApp");

class Config {
    public const int MaxRetries = 3;
    public const string AppName = "MyApp";
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
