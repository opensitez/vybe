// vybe-test: csharp/csharp_lambda_expression_typing/lambda_expression_typing_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expression_typing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// lambda_expression_typing
var map = new System.Collections.Generic.Dictionary<int, int>(); map[76] = 77; __Check((map.ContainsKey(76) && map[76] == 77).ToString(), "True");
