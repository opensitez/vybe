// vybe-test: csharp/csharp_structs_value_semantics/struct_can_be_stored_inside_array
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

var points = new[] { new Point { X = 3 }, new Point { X = 4 } }
;
foreach (var point in points) __P((point.X).ToString());
__Check("3\n4");

struct Point { public int X; }

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
