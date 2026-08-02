// vybe-test: csharp/csharp_enum_guid_version/enum_get_underlying_type_reports_int_by_default
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum State { Idle } __Check((System.Enum.GetUnderlyingType(typeof(State)).Name).ToString(), "Int32");
