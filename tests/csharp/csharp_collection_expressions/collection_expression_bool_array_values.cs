// vybe-test: csharp/csharp_collection_expressions/collection_expression_bool_array_values
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool[] flags = [true, false, true];
__Check((flags[0]).ToString(), "True"); __Check((flags[1]).ToString(), "False");
