// vybe-test: csharp/csharp_collection_expressions/collection_expression_short_array
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

short[] vals = [10, 20];
__Check((vals[1]).ToString(), "20");
