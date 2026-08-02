// vybe-test: csharp/csharp_collection_expressions/collection_expression_char_array_elements
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char[] chars = ['x', 'y'];
__Check((chars[0]).ToString(), "x");
