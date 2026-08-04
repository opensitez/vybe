// vybe-test: csharp/csharp_nameof_expressions/nameof_of_method_group_in_expression_context
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

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

class Ops{public void Execute(){}} __P((nameof(Ops)+"."+nameof(Ops.Execute)).ToString());
__Check("Ops.Execute");
