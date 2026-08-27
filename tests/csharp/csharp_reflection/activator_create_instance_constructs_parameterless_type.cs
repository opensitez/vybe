// vybe-test: csharp/csharp_reflection/activator_create_instance_constructs_parameterless_type
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

using static __Harness;

var w = (Widget)System.Activator.CreateInstance(typeof(Widget));
__P((w.Value).ToString());
__Check("42");

class Widget { public int Value = 42; }

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
