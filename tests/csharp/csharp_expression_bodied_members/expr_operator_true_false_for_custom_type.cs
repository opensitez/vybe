// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_true_false_for_custom_type
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

Flag f = new Flag { On = true }
;
if (f) __P(("yes").ToString());
else __P(("no").ToString());
__Check("yes");

struct Flag { public bool On; public static bool operator true(Flag f) => f.On; public static bool operator false(Flag f) => !f.On; }

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
