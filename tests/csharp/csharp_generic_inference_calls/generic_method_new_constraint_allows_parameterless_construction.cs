// vybe-test: csharp/csharp_generic_inference_calls/generic_method_new_constraint_allows_parameterless_construction
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

using static __Harness;

T Create<T>() where T : new() { return new T(); }
__P((Create<Widget>().Size).ToString());
__Check("4");

class Widget { public int Size = 4; }

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
