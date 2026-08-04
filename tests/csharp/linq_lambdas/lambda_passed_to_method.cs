// vybe-test: csharp/linq_lambdas/lambda_passed_to_method
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

class Processor {
    public int Apply(int value, Func<int, int> transform) {
        return transform(value);
    }
}
var p = new Processor();
__P((p.Apply(5, x => x * x)).ToString());
__P((p.Apply(5, x => x + 10)).ToString());
__Check("25\n15");
