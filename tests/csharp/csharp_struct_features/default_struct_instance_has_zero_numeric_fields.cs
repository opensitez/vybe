// vybe-test: csharp/csharp_struct_features/default_struct_instance_has_zero_numeric_fields
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

using static __Harness;

Size s = default;
__P((s.W).ToString());
__P((s.H).ToString());
__Check("0\n0");

struct Size { public int W, H; }

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
