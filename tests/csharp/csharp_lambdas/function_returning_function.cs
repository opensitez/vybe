// vybe-test: csharp/csharp_lambdas/function_returning_function
// origin: languages/csharp/tests/csharp/test_csharp_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Func<int, int> Multiplier(int factor) {
    return x => x * factor;
}
var triple = Multiplier(3);
__Check((triple(7)).ToString(), "21");
