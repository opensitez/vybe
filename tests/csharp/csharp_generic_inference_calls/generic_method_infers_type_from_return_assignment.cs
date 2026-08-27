// vybe-test: csharp/csharp_generic_inference_calls/generic_method_infers_type_from_return_assignment
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

using static __Harness;

T First<T>(T left, T right) { return left; }
string chosen = First("left", "right");
__P((chosen).ToString());
__Check("left");

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
