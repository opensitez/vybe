// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_serializable_plain_type_false
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

using static __Harness;

App.Run();
__Check("Run_attribute_serializable_plain_type_false");

class App {
    public static void Run() => __P("Run_attribute_serializable_plain_type_false");
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
