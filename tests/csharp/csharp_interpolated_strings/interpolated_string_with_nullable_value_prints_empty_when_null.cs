// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_with_nullable_value_prints_empty_when_null
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? value = null; __Check(($"[{value}]").ToString(), "[]");
