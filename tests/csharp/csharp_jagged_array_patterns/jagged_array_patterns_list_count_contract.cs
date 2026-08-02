// vybe-test: csharp/csharp_jagged_array_patterns/jagged_array_patterns_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_jagged_array_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// jagged_array_patterns
var values = new System.Collections.Generic.List<int> { 28, 29, 28 }; __Check((values.Count == 3).ToString(), "True");
