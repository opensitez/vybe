// vybe-test: csharp/csharp_oop/interface_with_multiple_methods
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ICalculator {
    int Add(int a, int b);
    int Multiply(int a, int b);
}
class Calc : ICalculator {
    public int Add(int a, int b) { return a + b; }
    public int Multiply(int a, int b) { return a * b; }
}
var c = new Calc();
__Check((c.Add(3, 4)).ToString(), "7");
__Check((c.Multiply(3, 4)).ToString(), "12");
