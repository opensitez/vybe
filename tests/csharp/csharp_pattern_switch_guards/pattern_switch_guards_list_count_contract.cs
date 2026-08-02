// vybe-test: csharp/csharp_pattern_switch_guards/pattern_switch_guards_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_guards.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_switch_guards
var values = new System.Collections.Generic.List<int> { 42, 43, 42 }; __Check((values.Count == 3).ToString(), "True");
