// vybe-test: csharp/csharp_nameof_expressions/nameof_in_interpolated_string_embeds_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string title="demo"; __Check(($"name={nameof(title)}").ToString(), "name=title");
