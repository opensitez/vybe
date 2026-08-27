// vybe-test: csharp/csharp_collection_initializer_syntax/object_initializer_sets_public_fields_before_use
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

using static __Harness;

var point = new Point { X = 2, Y = 5 }
;
__P((point.X + point.Y).ToString());
__Check("7");

class Point { public int X; public int Y; }

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
