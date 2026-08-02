// vybe-test: csharp/csharp_guid_parse_format/guid_to_string_with_format_specifier_renders_hyphenated_value
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_format.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var id = System.Guid.Parse("11111111-2222-3333-4444-555555555555");
__Check((id.ToString("D").StartsWith("11111111")).ToString(), "True");
