// vybe-test: csharp/csharp_math_advanced/math_round_midpoint_away_from_zero
// origin: languages/csharp/tests/csharp/test_csharp_math_advanced.rs

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

__P((System.Math.Round(2.5,System.MidpointRounding.AwayFromZero)).ToString());
__Check("3");
