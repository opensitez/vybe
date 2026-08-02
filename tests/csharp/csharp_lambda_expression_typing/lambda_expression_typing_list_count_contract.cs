// vybe-test: csharp/csharp_lambda_expression_typing/lambda_expression_typing_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expression_typing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_expression_typing
var values = new System.Collections.Generic.List<int> { 76, 77, 76 }; __Check((values.Count == 3).ToString(), "True");
