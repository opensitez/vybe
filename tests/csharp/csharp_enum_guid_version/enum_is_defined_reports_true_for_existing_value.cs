// vybe-test: csharp/csharp_enum_guid_version/enum_is_defined_reports_true_for_existing_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum State { Idle, Running } __Check((System.Enum.IsDefined(typeof(State), "Running")).ToString(), "True");
