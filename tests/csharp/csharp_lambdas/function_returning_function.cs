// vybe-test: csharp/csharp_lambdas/function_returning_function
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

Func<int, int> Multiplier(int factor) {
    return x => x * factor;
}
var triple = Multiplier(3);
__P((triple(7)).ToString());
__Check("21");
