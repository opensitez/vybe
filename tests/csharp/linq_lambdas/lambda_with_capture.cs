// vybe-test: csharp/linq_lambdas/lambda_with_capture
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

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

int multiplier = 3;
Func<int, int> mul = x => x * multiplier;
__P((mul(10)).ToString());
__P((mul(7)).ToString());
__Check("30\n21");
