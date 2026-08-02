// vybe-test: csharp/csharp_record_shape_basics/record_shape_basics_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_record_shape_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// record_shape_basics
var values = new System.Collections.Generic.List<int> { 39, 40, 39 }; __Check((values.Count == 3).ToString(), "True");
