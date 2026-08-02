// vybe-test: csharp/csharp_enum_guid_version/enum_format_d_outputs_numeric_representation
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum State { Idle = 1 } __Check((System.Enum.Format(typeof(State), State.Idle, "D")).ToString(), "1");
