// vybe-test: csharp/csharp_operators/compound_assignment
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

int x = 10;
x += 5; __P((x).ToString());
x -= 3; __P((x).ToString());
x *= 2; __P((x).ToString());
x /= 4; __P((x).ToString());
x %= 5; __P((x).ToString());
__Check("15\n12\n24\n6\n1");
