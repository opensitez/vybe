// vybe-test: csharp/csharp_numeric_precision/decimal_avoids_float_rounding_error
// origin: languages/csharp/tests/csharp/test_csharp_numeric_precision.rs

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

decimal a=0.1m, b=0.2m;
__P((a+b==0.3m).ToString());
__Check("True");
