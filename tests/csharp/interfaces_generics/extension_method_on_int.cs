// vybe-test: csharp/interfaces_generics/extension_method_on_int
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class IntExtensions {
    public static bool IsEven(this int n) { return n % 2 == 0; }
    public static int Square(this int n) { return n * n; }
}
__Check((4.IsEven()).ToString(), "True");
__Check((3.IsEven()).ToString(), "False");
__Check((5.Square()).ToString(), "25");
