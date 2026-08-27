// vybe-test: csharp/csharp_reflection/field_info_get_value_reads_public_field
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

using static __Harness;

var obj = new Data();
var field = typeof(Data).GetField("X");
__P((field.GetValue(obj)).ToString());
__Check("3");

class Data { public int X = 3; }

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
