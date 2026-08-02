// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_in_return_expression_from_local_function
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Label(int n) { return $"n={n}"; }
__Check((Label(5)).ToString(), "n=5");
