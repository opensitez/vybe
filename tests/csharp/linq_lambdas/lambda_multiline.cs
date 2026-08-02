// vybe-test: csharp/linq_lambdas/lambda_multiline
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Func<int, int> factorial = null;
factorial = n => {
    if (n <= 1) return 1;
    return n * factorial(n - 1);
};
__Check((factorial(5)).ToString(), "120");
