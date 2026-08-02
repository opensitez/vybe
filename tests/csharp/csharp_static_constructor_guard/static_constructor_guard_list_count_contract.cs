// vybe-test: csharp/csharp_static_constructor_guard/static_constructor_guard_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_guard.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// static_constructor_guard
var values = new System.Collections.Generic.List<int> { 69, 70, 69 }; __Check((values.Count == 3).ToString(), "True");
