// vybe-test: csharp/csharp_enum_guid_version/version_parse_exposes_major_minor_and_build
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var version = System.Version.Parse("2.4.6"); __Check((version.Major).ToString(), "2"); __Check((version.Minor).ToString(), "4"); __Check((version.Build).ToString(), "6");
