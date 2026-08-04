// vybe-test: csharp/csharp_lambda_expressions/lambda_passed_as_argument_to_higher_order_method
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

int Apply(System.Func<int,int,int> op, int a, int b) => op(a,b);
__P((Apply((a,b) => a+b, 3, 4)).ToString());
__Check("7");
