// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_embeds_local_variable_text
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var name = "Ada"; __Check(($"{name}").ToString(), "Ada");
