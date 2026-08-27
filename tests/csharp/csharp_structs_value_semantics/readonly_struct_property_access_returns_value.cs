// vybe-test: csharp/csharp_structs_value_semantics/readonly_struct_property_access_returns_value
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

__P((new Size(6).Width).ToString());
__Check("6");

readonly struct Size { public int Width { get; } public Size(int width) { Width = width; } }

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
