// vybe-test: csharp/csharp_pattern_list/is_list_collection_expression_literal_pair
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data=[10,20]; __Check((data is [10,20]).ToString(), "True");
