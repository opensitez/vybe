// vybe-test: csharp/csharp_array_reverse_patterns/array_reverse_patterns_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_array_reverse_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_reverse_patterns
var values = new System.Collections.Generic.List<int> { 27, 28, 27 }; __Check((values.Count == 3).ToString(), "True");
