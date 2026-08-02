// vybe-test: csharp/csharp_enum_guid_version/version_to_string_round_trips_original_text
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var version = new System.Version(1, 2, 3, 4); __Check((version.ToString()).ToString(), "1.2.3.4");
