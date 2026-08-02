// vybe-test: csharp/csharp_lambdas/lambda_as_callback
// origin: languages/csharp/tests/csharp/test_csharp_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Apply(Func<int, int> fn, int x) {
    return fn(x);
}
__Check((Apply(x => x * x, 5)).ToString(), "25");
