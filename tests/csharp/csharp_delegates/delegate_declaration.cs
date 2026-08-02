// vybe-test: csharp/csharp_delegates/delegate_declaration
// origin: languages/csharp/tests/csharp/test_csharp_delegates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

delegate int MathOp(int a, int b);
class Program {
    public static int Add(int a, int b) { return a + b; }
    public static int Mul(int a, int b) { return a * b; }
}
MathOp op = (a, b) => a + b;
__Check((op(3, 4)).ToString(), "7");
op = (a, b) => a * b;
__Check((op(3, 4)).ToString(), "12");
