// vybe-test: csharp/csharp_string_builder_usage/string_builder_usage_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_string_builder_usage.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_builder_usage
var values = new System.Collections.Generic.List<int> { 20, 21, 20 }; __Check((values.Count == 3).ToString(), "True");
