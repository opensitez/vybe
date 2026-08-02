// vybe-test: csharp/csharp_collection_expressions/collection_expression_float_array
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

float[] vals = [1.1f, 2.2f];
__Check((vals.Length).ToString(), "2");
