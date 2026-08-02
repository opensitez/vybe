// vybe-test: csharp/linq_lambdas/lambda_with_capture
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int multiplier = 3;
Func<int, int> mul = x => x * multiplier;
__Check((mul(10)).ToString(), "30");
__Check((mul(7)).ToString(), "21");
