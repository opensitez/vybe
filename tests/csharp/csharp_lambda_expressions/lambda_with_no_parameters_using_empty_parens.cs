// vybe-test: csharp/csharp_lambda_expressions/lambda_with_no_parameters_using_empty_parens
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<string> greeting = () => "hello";
__Check((greeting()).ToString(), "hello");
