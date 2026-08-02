// vybe-test: csharp/csharp_guid_parse_format/guid_parse_accepts_standard_hyphenated_representation
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_format.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var id = System.Guid.Parse("11111111-2222-3333-4444-555555555555");
__Check((id.ToString().StartsWith("11111111")).ToString(), "True");
