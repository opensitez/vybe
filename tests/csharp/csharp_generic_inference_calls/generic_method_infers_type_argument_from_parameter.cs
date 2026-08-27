// vybe-test: csharp/csharp_generic_inference_calls/generic_method_infers_type_argument_from_parameter
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

using static __Harness;

T Identity<T>(T value) { return value; }
__P((Identity(42)).ToString());
__P((Identity("text")).ToString());
__Check("42\ntext");

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
