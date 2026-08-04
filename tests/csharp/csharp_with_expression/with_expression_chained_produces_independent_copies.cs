// vybe-test: csharp/csharp_with_expression/with_expression_chained_produces_independent_copies
// origin: languages/csharp/tests/csharp/test_csharp_with_expression.rs

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

record Box(int Width, int Height);
var a = new Box(1, 1);
var b = a with { Width = 2 };
var c = b with { Height = 3 };
__P((a.Width).ToString());
__P((c.Width).ToString());
__P((c.Height).ToString());
__Check("1\n2\n3");
