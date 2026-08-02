// vybe-test: csharp/csharp_for_loop_bounds/for_loop_bounds_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_for_loop_bounds.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// for_loop_bounds
var values = new System.Collections.Generic.List<int> { 45, 46, 45 }; __Check((values.Count == 3).ToString(), "True");
