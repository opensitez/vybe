// vybe-test: csharp/csharp_enum_guid_version/enum_try_parse_can_ignore_case_when_requested
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum State { Idle, Running } System.Enum.TryParse<State>("running", true, out var value); __Check((value).ToString(), "Running");
