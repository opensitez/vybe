// vybe-test: csharp/csharp_classes/recursive_factorial
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Factorial(int n) {
    if (n <= 1) return 1;
    return n * Factorial(n - 1);
}
__Check((Factorial(6)).ToString(), "720");
