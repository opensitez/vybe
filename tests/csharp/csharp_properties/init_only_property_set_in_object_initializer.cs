// vybe-test: csharp/csharp_properties/init_only_property_set_in_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

using static __Harness;

var p = new Point { X=1, Y=2 }
;
__P((p.X).ToString());
__P((p.Y).ToString());
__Check("1\n2");

class Point { public int X { get; init; } public int Y { get; init; } }

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
