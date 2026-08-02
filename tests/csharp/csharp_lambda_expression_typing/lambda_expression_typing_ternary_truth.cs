// vybe-test: csharp/csharp_lambda_expression_typing/lambda_expression_typing_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expression_typing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_expression_typing
int seed = 76; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
