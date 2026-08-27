// vybe-test: csharp/csharp_classes/class_auto_property
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

using static __Harness;

var c = new Config();
c.Name = "test";
c.Value = 42;
__P((c.Name).ToString());
__P((c.Value).ToString());
__Check("test\n42");

class Config {
    public string Name { get; set; }
    public int Value { get; set; }
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
