// vybe-test: csharp/csharp_lambda_expression_typing/lambda_expression_typing_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expression_typing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_expression_typing
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
