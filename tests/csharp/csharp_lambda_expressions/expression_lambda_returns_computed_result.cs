// vybe-test: csharp/csharp_lambda_expressions/expression_lambda_returns_computed_result
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

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

System.Func<int,int> f = x => x*x;
__P((f(5)).ToString());
__Check("25");
