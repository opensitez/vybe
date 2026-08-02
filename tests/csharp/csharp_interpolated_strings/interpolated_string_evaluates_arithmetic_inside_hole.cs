// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_evaluates_arithmetic_inside_hole
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int a = 6; int b = 7; __Check(($"{a + b}").ToString(), "13");
