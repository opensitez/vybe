// vybe-test: csharp/csharp_property_accessor_matrix/property_accessor_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_property_accessor_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// property_accessor_matrix
var values = new System.Collections.Generic.List<int> { 64, 65, 64 }; __Check((values.Count == 3).ToString(), "True");
