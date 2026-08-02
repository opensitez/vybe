// vybe-test: csharp/csharp_array_copy_behavior/array_copy_behavior_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_array_copy_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_copy_behavior
var values = new System.Collections.Generic.List<int> { 26, 27, 26 }; __Check((values.Count == 3).ToString(), "True");
