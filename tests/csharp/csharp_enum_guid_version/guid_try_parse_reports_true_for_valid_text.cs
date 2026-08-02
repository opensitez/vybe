// vybe-test: csharp/csharp_enum_guid_version/guid_try_parse_reports_true_for_valid_text
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var ok = System.Guid.TryParse("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee", out var value); __Check((ok).ToString(), "True"); __Check((value.ToString()).ToString(), "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee");
