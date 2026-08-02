// vybe-test: csharp/csharp_string_split_join/string_split_join_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_split_join
var values = new System.Collections.Generic.List<int> { 21, 22, 21 }; __Check((values.Count == 3).ToString(), "True");
