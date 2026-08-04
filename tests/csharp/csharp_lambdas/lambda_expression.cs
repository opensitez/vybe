// vybe-test: csharp/csharp_lambdas/lambda_expression
// origin: languages/csharp/tests/csharp/test_csharp_lambdas.rs

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

var double_it = (int x) => x * 2;
__P((double_it(5)).ToString());
__Check("10");
