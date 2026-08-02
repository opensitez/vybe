// vybe-test: csharp/csharp_lambda_expression_typing/lambda_expression_typing_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expression_typing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_expression_typing
var tuple = (left: 76, right: 77); __Check((tuple.left < tuple.right).ToString(), "True");
