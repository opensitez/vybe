// vybe-test: csharp/csharp_collection_initializer_syntax/array_initializer_literal_sets_length_and_elements
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var data = new[] { 10, 20, 30 };
__Check((data.Length).ToString(), "3");
__Check((data[1]).ToString(), "20");
