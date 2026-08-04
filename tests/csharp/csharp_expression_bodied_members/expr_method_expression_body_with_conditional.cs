// vybe-test: csharp/csharp_expression_bodied_members/expr_method_expression_body_with_conditional
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

class Sign { public string Label(int n) => n < 0 ? "neg" : n > 0 ? "pos" : "zero"; }
__P((new Sign().Label(-1)).ToString()); __P((new Sign().Label(0)).ToString()); __P((new Sign().Label(2)).ToString());
__Check("neg\nzero\npos");
