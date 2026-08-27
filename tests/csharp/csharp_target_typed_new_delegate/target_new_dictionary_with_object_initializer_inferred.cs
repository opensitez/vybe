// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_dictionary_with_object_initializer_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

using static __Harness;

System.Collections.Generic.Dictionary<string, int> map = new() { ["a"] = 1, ["b"] = 2 }
;
__P((map["b"]).ToString());
__Check("2");

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
