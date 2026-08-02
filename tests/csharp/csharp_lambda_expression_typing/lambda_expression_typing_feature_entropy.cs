// vybe-test: csharp/csharp_lambda_expression_typing/lambda_expression_typing_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expression_typing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_expression_typing
string feature = "lambda_expression_typing:76"; __Check((feature.Length >= 1).ToString(), "True");
