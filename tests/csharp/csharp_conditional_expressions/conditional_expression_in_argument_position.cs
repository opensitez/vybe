// vybe-test: csharp/csharp_conditional_expressions/conditional_expression_in_argument_position
// origin: languages/csharp/tests/csharp/test_csharp_conditional_expressions.rs

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

int n=7;
__P((string.Format("{0}",n%2==0?"even":"odd")).ToString());
__Check("odd");
