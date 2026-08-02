// vybe-test: csharp/csharp_array_apis/array_true_for_all_checks_entire_sequence
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var values = new[] { 2, 4, 6 }; __Check((System.Array.TrueForAll(values, value => value % 2 == 0)).ToString(), "True");
