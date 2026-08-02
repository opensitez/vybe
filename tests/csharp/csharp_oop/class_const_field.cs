// vybe-test: csharp/csharp_oop/class_const_field
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class MathConst {
    public const double PI = 3.14159;
    public const double E = 2.71828;
}
__Check((MathConst.PI).ToString(), "3.14159");
__Check((MathConst.E).ToString(), "2.71828");
