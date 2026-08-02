// vybe-test: csharp/csharp_collection_expressions/collection_expression_string_join
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string[] parts = ["a", "b", "c"];
__Check((string.Join("-", parts)).ToString(), "a-b-c");
