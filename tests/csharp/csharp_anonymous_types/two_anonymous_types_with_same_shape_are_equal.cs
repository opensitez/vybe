// vybe-test: csharp/csharp_anonymous_types/two_anonymous_types_with_same_shape_are_equal
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_types.rs

using static __Harness;

var a=new{X=1,Y=2}
;
var b=new{X=1,Y=2}
;
__P((a.Equals(b)).ToString());
__Check("True");

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
