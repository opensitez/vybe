// vybe-test: csharp/csharp_lambdas/lambda_closure
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

int counter = 0;
var inc = () => { counter++; };
inc();
inc();
inc();
__P((counter).ToString());
__Check("3");
