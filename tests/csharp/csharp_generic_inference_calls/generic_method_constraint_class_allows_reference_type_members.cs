// vybe-test: csharp/csharp_generic_inference_calls/generic_method_constraint_class_allows_reference_type_members
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

using static __Harness;

__P((Describe("data")).ToString());
__Check("data");

string Describe<T>(T value) where T : class {
    return value == null ? "null" : value.ToString();
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
