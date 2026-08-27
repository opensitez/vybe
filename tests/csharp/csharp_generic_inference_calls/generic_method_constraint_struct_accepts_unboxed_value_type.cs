// vybe-test: csharp/csharp_generic_inference_calls/generic_method_constraint_struct_accepts_unboxed_value_type
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

using static __Harness;

__P((Scale(6)).ToString());
__Check("12");

int Scale<T>(T value) where T : struct {
    return 2 * (int)(object)value;
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
