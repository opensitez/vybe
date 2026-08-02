// vybe-test: csharp/oop_advanced/static_class_methods
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class MathHelper {
    public static int Square(int x) { return x * x; }
    public static int Double(int x) { return x * 2; }
}
__Check((MathHelper.Square(5)).ToString(), "25");
__Check((MathHelper.Double(7)).ToString(), "14");
