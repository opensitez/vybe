// vybe-test: csharp/csharp_lambdas/lambda_expression
// origin: languages/csharp/tests/csharp/test_csharp_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var double_it = (int x) => x * 2;
__Check((double_it(5)).ToString(), "10");
