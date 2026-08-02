// vybe-test: csharp/csharp_collection_expressions/collection_expression_string_array_elements
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string[] words = ["a", "b", "c"];
__Check((words[1]).ToString(), "b");
