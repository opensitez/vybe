// vybe-test: csharp/control_flow/switch_basic
// origin: languages/csharp/tests/csharp/test_control_flow.rs

using static __Harness;

var x = 2;
var result = "";
switch (x) {
            case 1: result = "one"; break;
            case 2: result = "two"; break;
            case 3: result = "three"; break;
            default: result = "other"; break;
        }
__P((result).ToString());
__Check("two");

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
