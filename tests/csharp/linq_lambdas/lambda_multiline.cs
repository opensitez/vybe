// vybe-test: csharp/linq_lambdas/lambda_multiline
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

Func<int, int> factorial = null;
factorial = n => {
    if (n <= 1) return 1;
    return n * factorial(n - 1);
};
__P((factorial(5)).ToString());
__Check("120");
