// vybe-test: csharp/csharp_math_functions/math_truncate_removes_fractional_part_toward_zero
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

__P((System.Math.Truncate(9.9)).ToString());
__Check("9");
