// vybe-test: csharp/csharp_math_functions/math_pow_raises_base_to_integer_exponent
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

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

__P((System.Math.Pow(2, 10)).ToString());
__Check("1024");
