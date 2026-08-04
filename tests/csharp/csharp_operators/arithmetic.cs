// vybe-test: csharp/csharp_operators/arithmetic
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

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

__P((10 + 5).ToString());
__P((10 - 5).ToString());
__P((10 * 5).ToString());
__P((10 % 3).ToString());
__Check("15\n5\n50\n1");
