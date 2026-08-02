// vybe-test: csharp/csharp_method_overload_resolution/optional_parameter_uses_default_when_argument_omitted
// origin: languages/csharp/tests/csharp/test_csharp_method_overload_resolution.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string FormatLine(string text, int level = 1) {
    return level + ":" + text;
}
__Check((FormatLine("ok")).ToString(), "1:ok");
__Check((FormatLine("warn", 3)).ToString(), "3:warn");
