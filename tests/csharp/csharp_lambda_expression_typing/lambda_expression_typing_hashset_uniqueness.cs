// vybe-test: csharp/csharp_lambda_expression_typing/lambda_expression_typing_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expression_typing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_expression_typing
var set = new System.Collections.Generic.HashSet<int>(); set.Add(76); set.Add(76); __Check((set.Count == 1).ToString(), "True");
