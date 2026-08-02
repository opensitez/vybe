// vybe-test: csharp/csharp_lambda_expression_typing/lambda_expression_typing_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expression_typing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_expression_typing
double seed = 76; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
