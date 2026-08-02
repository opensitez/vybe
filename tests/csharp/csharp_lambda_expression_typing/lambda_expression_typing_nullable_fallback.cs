// vybe-test: csharp/csharp_lambda_expression_typing/lambda_expression_typing_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expression_typing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_expression_typing
int? maybe = null; int fallback = maybe ?? 76; __Check((fallback == 76).ToString(), "True");
