// vybe-test: csharp/csharp_enum_guid_version/guid_try_parse_reports_false_for_invalid_text
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var ok = System.Guid.TryParse("bad-guid", out var value); __Check((ok).ToString(), "False");
