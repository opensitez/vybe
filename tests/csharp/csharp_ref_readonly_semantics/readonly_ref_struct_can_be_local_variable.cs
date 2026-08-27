// vybe-test: csharp/csharp_ref_readonly_semantics/readonly_ref_struct_can_be_local_variable
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

using static __Harness;

int val = 42;
ref readonly int r = ref val;
__P(r.ToString());
__Check("42");
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
