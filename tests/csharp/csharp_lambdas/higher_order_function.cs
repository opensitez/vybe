// vybe-test: csharp/csharp_lambdas/higher_order_function
// origin: languages/csharp/tests/csharp/test_csharp_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Func<int, int> square = x => x * x;
Func<int, int> negate = x => -x;
__Check((square(5)).ToString(), "25");
__Check((negate(5)).ToString(), "-5");
