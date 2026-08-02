// vybe-test: csharp/csharp_enum_guid_version/version_compare_to_reports_negative_for_smaller_version
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var left = new System.Version(1, 2); var right = new System.Version(1, 3); __Check((left.CompareTo(right)).ToString(), "-1");
