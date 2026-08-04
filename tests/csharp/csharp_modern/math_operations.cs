// vybe-test: csharp/csharp_modern/math_operations
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

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

__P((Math.Abs(-42)).ToString());
__P((Math.Max(10, 20)).ToString());
__P((Math.Min(10, 20)).ToString());
__P((Math.Sqrt(25)).ToString());
__P((Math.Floor(3.7)).ToString());
__P((Math.Ceiling(3.2)).ToString());
__Check("42\n20\n10\n5\n3\n4");
