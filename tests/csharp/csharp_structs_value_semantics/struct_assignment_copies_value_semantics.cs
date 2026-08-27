// vybe-test: csharp/csharp_structs_value_semantics/struct_assignment_copies_value_semantics
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

var left = new Counter { Value = 1 }
;
var right = left;
right.Value = 9;
__P((left.Value).ToString());
__P((right.Value).ToString());
__Check("1\n9");

struct Counter { public int Value; }

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
