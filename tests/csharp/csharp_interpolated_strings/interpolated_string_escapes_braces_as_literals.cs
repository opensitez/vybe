// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_escapes_braces_as_literals
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 3; __Check(($"{{count}}={n}").ToString(), "{count}=3");
