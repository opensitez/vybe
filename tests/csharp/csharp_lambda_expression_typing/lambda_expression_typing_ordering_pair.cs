// vybe-test: csharp/csharp_lambda_expression_typing/lambda_expression_typing_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expression_typing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_expression_typing
int seed = 76; int right = seed + 1; __Check((seed < right).ToString(), "True");
