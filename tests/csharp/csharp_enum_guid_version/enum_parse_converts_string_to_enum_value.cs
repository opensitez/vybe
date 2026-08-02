// vybe-test: csharp/csharp_enum_guid_version/enum_parse_converts_string_to_enum_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum State { Idle, Running, Done } __Check((System.Enum.Parse(typeof(State), "Running")).ToString(), "Running");
