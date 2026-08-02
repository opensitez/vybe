// vybe-test: csharp/common_patterns/factorial_recursive
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Math2 {
    public static int Factorial(int n) {
        if (n <= 1) return 1;
        return n * Factorial(n - 1);
    }
}
__Check((Math2.Factorial(0)).ToString(), "1");
__Check((Math2.Factorial(1)).ToString(), "1");
__Check((Math2.Factorial(5)).ToString(), "120");
__Check((Math2.Factorial(10)).ToString(), "3628800");
