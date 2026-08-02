// vybe-test: csharp/csharp_collection_expressions/collection_expression_byte_array_length
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] data = [1, 2, 3, 4];
__Check((data.Length).ToString(), "4");
