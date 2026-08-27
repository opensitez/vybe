// vybe-test: csharp/csharp_generic_inference_calls/generic_method_overload_resolution_prefers_specific_argument_types
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

using static __Harness;

__P(Picker.Pick("Hello"));
__Check("Hello");

class Picker {
    public static string Pick(string s) => s;
    public static int Pick(int i) => i;
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
