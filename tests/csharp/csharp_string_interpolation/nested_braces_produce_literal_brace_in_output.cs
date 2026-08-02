// vybe-test: csharp/csharp_string_interpolation/nested_braces_produce_literal_brace_in_output
// origin: languages/csharp/tests/csharp/test_csharp_string_interpolation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=5; __Check(($"{{n}}={n}").ToString(), "{n}=5");
