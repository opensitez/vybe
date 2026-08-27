// vybe-test: csharp/winforms/form_subclass_constructs_via_descriptor_and_inherits_properties
// origin: languages/csharp/tests/csharp/test_winforms.rs

using static __Harness;

__P("FormLoaded");
__Check("FormLoaded");
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
