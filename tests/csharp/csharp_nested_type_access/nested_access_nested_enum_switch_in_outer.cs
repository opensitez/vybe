// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_enum_switch_in_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Gate().Label(Gate.Mode.On)).ToString());
__Check("on");

class Gate{public enum Mode{On,Off} public string Label(Mode m){switch(m){case Mode.On:return "on"; default:return "off";}}}

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
