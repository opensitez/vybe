// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_true_false_for_custom_type
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

struct Flag { public bool On; public static bool operator true(Flag f) => f.On; public static bool operator false(Flag f) => !f.On; }
Flag f = new Flag { On = true }; if (f) __P(("yes").ToString()); else __P(("no").ToString());
__Check("yes");
