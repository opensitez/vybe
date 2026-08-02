// vybe-test: csharp/csharp_collection_expressions/collection_expression_decimal_array
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal[] vals = [1.5m, 2.5m];
__Check((vals[0] + vals[1]).ToString(), "4.0");
