// vybe-test: csharp/csharp_type_conversions/string_interpolation_converts_integer_to_text
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int count = 9; __Check(($"count={count}").ToString(), "count=9");
