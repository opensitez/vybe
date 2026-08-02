// vybe-test: csharp/csharp_array_length_variants/array_length_variants_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_array_length_variants.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_length_variants
var values = new System.Collections.Generic.List<int> { 25, 26, 25 }; __Check((values.Count == 3).ToString(), "True");
