// vybe-test: csharp/csharp_patterns/nested_class_access
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

using static __Harness;

var o = new Outer();
var i = new Outer.Inner();
__P((o.Value).ToString());
__P((i.Value).ToString());
__Check("10\n20");

class Outer {
    public int Value = 10;
    public class Inner {
        public int Value = 20;
    }
}

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
