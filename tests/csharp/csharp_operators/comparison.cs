// vybe-test: csharp/csharp_operators/comparison
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

__P((1 < 2).ToString());
__P((2 > 1).ToString());
__P((1 <= 1).ToString());
__P((1 >= 1).ToString());
__P((1 == 1).ToString());
__P((1 != 2).ToString());
__Check("True\nTrue\nTrue\nTrue\nTrue\nTrue");
