// vybe-test: csharp/csharp_checked_context_math/checked_context_math_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_checked_context_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// checked_context_math
var values = new System.Collections.Generic.List<int> { 12, 13, 12 }; __Check((values.Count == 3).ToString(), "True");
