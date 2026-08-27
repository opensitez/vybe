// vybe-test: csharp/csharp_properties/auto_property_with_default_value_initializer
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

using static __Harness;

__P((new Config().Timeout).ToString());
__Check("30");

class Config { public int Timeout { get; set; } = 30; }

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
