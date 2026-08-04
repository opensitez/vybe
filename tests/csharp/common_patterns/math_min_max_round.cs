// vybe-test: csharp/common_patterns/math_min_max_round
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

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

__P((Math.Min(3, 7)).ToString());
__P((Math.Max(3, 7)).ToString());
__P((Math.Round(3.7)).ToString());
__P((Math.Floor(3.7)).ToString());
__P((Math.Ceiling(3.2)).ToString());
__Check("3\n7\n4\n3\n4");
