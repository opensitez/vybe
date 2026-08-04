// vybe-test: csharp/csharp_numeric_ops/double_division_by_zero_produces_infinity
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

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

double d=1.0/0.0;
__P((double.IsInfinity(d)).ToString());
__Check("True");
