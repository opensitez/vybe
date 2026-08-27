// vybe-test: csharp/csharp_reflection_emit/field_info_get_and_set_value_on_instance
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

using static __Harness;

var fi=typeof(Box).GetField("V");
var obj=new Box();
fi.SetValue(obj,55);
__P((fi.GetValue(obj)).ToString());
__Check("55");

class Box{public int V;}

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
