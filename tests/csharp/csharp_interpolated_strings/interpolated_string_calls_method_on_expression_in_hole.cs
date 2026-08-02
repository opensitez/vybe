// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_calls_method_on_expression_in_hole
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var text = "hi"; __Check(($"{text.ToUpper()}").ToString(), "HI");
