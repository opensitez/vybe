// vybe-test: csharp/math/math_div_rem_returns_quotient_and_remainder
// origin: languages/csharp/tests/csharp/test_math.rs

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

int remainder;
var quotient = System.Math.DivRem(17, 5, out remainder);
__P((quotient).ToString());
__P((remainder).ToString());
__Check("3\n2");
