// vybe-test: csharp/csharp_reflection_emit/property_info_get_set_accessor_names
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

using static __Harness;

var pi=typeof(Model).GetProperty("Value");
__P((pi.CanRead).ToString());
__P((pi.CanWrite).ToString());
__Check("True\nTrue");

class Model{public int Value{get;set;}}

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
