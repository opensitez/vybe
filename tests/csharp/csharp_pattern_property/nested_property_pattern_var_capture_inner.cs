// vybe-test: csharp/csharp_pattern_property/nested_property_pattern_var_capture_inner
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

object o=new Outer{Child=new Inner{N=9}}
;
if(o is Outer{Child:{N:var n}}) __P((n).ToString());
__Check("9");

class Inner { public int N; }

class Outer { public Inner Child; }

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
