// vybe-test: csharp/oop_advanced/sealed_class_basic
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

var c = new Config("prod");
__P((c.Name).ToString());
__Check("prod");

sealed class Config {
    public string Name { get; set; }
    public Config(string n) { Name = n; }
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
