// vybe-test: csharp/csharp_abstract_sealed/sealed_class_cannot_be_used_as_base_detected_at_compile_time_but_runtime_ok
// origin: languages/csharp/tests/csharp/test_csharp_abstract_sealed.rs

using static __Harness;

var f = new Final();
__P((f.Value).ToString());
__Check("7");

sealed class Final { public int Value = 7; }

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
