// vybe-test: csharp/csharp_static_type_behaviors/instance_method_calls_static_helper_on_same_type
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

using static __Harness;

var converter = new Converter();
__P((converter.Convert(5)).ToString());
__Check("11");

class Converter {
    public static int Double(int value) { return value * 2; }
    public int Convert(int value) { return Double(value) + 1; }
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
