// vybe-test: csharp/csharp_oop/class_auto_property_default
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

using static __Harness;

var c = new Config();
__P((c.Name).ToString());
c.Name = "custom";
__P((c.Name).ToString());
__Check("default\ncustom");

class Config {
    public string Name { get; set; } = "default";
    public int Count { get; set; } = 0;
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
