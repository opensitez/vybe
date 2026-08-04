// vybe-test: csharp/math/math_clamp_restricts_value_to_inclusive_bounds
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

__P((System.Math.Clamp(10, 0, 5)).ToString());
__Check("5");
