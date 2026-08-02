// vybe-test: csharp/csharp_enum_guid_version/version_equals_reports_true_for_identical_versions
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((new System.Version(3, 5).Equals(new System.Version(3, 5))).ToString(), "True");
