// vybe-test: csharp/csharp_lambda_expression_typing/lambda_expression_typing_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expression_typing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_expression_typing
string feature = "lambda_expression_typing"; __Check((feature[0] == feature[0]).ToString(), "True");
