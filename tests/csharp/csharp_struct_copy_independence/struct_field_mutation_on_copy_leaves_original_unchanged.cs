// vybe-test: csharp/csharp_struct_copy_independence/struct_field_mutation_on_copy_leaves_original_unchanged
// origin: languages/csharp/tests/csharp/test_csharp_struct_copy_independence.rs

using static __Harness;

var left = new Point { X = 1 }
;
var right = left;
right.X = 9;
__P((left.X).ToString());
__P((right.X).ToString());
__Check("1\n9");

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
