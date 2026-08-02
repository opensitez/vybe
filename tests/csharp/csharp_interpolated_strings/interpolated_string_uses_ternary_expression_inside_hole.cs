// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_uses_ternary_expression_inside_hole
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 4; __Check(($"{(n % 2 == 0 ? "even" : "odd")}").ToString(), "even");
