// vybe-test: csharp/csharp_static_classes/static_readonly_field_set_in_static_constructor
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

using static __Harness;

__P((Config.Version).ToString());
__Check("1.0");

class Config {
    public static readonly string Version;
    static Config() { Version = "1.0"; }
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
