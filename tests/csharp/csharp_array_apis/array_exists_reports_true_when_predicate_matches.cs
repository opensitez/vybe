// vybe-test: csharp/csharp_array_apis/array_exists_reports_true_when_predicate_matches
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var values = new[] { 1, 3, 5 }; __Check((System.Array.Exists(values, value => value == 3)).ToString(), "True");
