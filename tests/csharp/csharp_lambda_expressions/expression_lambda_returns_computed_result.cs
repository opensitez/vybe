// vybe-test: csharp/csharp_lambda_expressions/expression_lambda_returns_computed_result
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int,int> f = x => x*x;
__Check((f(5)).ToString(), "25");
