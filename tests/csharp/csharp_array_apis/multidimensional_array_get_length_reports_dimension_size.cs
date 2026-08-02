// vybe-test: csharp/csharp_array_apis/multidimensional_array_get_length_reports_dimension_size
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var grid = new int[2, 3]; __Check((grid.GetLength(0)).ToString(), "2"); __Check((grid.GetLength(1)).ToString(), "3");
