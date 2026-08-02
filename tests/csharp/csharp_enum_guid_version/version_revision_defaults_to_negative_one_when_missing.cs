// vybe-test: csharp/csharp_enum_guid_version/version_revision_defaults_to_negative_one_when_missing
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var version = new System.Version(1, 2, 3); __Check((version.Revision).ToString(), "-1");
