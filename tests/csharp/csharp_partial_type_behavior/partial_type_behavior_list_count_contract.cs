// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// partial_type_behavior
var values = new System.Collections.Generic.List<int> { 70, 71, 70 }; __Check((values.Count == 3).ToString(), "True");
