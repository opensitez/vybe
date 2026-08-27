// vybe-test: csharp/csharp_pattern_property/is_property_pattern_var_capture_reads_field
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

object o=new Box{Value=25}
;
if(o is Box{Value:var v}) __P((v).ToString());
__Check("25");

class Box { public int Value; }

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
