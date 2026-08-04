// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_class_less_than
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

class Score { public int V; public static bool operator <(Score a, Score b) => a.V < b.V; public static bool operator >(Score a, Score b) => a.V > b.V; }
__P((new Score { V = 1 } < new Score { V = 2 }).ToString()); __P((new Score { V = 5 } > new Score { V = 3 }).ToString());
__Check("True\nTrue");
