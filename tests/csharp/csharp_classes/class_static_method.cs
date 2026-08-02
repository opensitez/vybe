// vybe-test: csharp/csharp_classes/class_static_method
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class MathUtils {
    public static int Square(int x) { return x * x; }
}
__Check((MathUtils.Square(7)).ToString(), "49");
