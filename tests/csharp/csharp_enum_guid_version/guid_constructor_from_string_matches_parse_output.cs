// vybe-test: csharp/csharp_enum_guid_version/guid_constructor_from_string_matches_parse_output
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var text = "11111111-2222-3333-4444-555555555555"; __Check((new System.Guid(text).ToString()).ToString(), "11111111-2222-3333-4444-555555555555");
