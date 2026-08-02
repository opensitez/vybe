// vybe-test: csharp/csharp_multidim_array_traversal/multidim_array_traversal_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_multidim_array_traversal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// multidim_array_traversal
var values = new System.Collections.Generic.List<int> { 29, 30, 29 }; __Check((values.Count == 3).ToString(), "True");
