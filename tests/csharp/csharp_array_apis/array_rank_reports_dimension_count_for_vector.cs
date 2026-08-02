// vybe-test: csharp/csharp_array_apis/array_rank_reports_dimension_count_for_vector
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var values = new[] { 1, 2, 3 }; __Check((values.Rank).ToString(), "1");
