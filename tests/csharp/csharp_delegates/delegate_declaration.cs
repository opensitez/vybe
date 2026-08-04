// vybe-test: csharp/csharp_delegates/delegate_declaration
// origin: languages/csharp/tests/csharp/test_csharp_delegates.rs

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

delegate int MathOp(int a, int b);
class Program {
    public static int Add(int a, int b) { return a + b; }
    public static int Mul(int a, int b) { return a * b; }
}
MathOp op = (a, b) => a + b;
__P((op(3, 4)).ToString());
op = (a, b) => a * b;
__P((op(3, 4)).ToString());
__Check("7\n12");
