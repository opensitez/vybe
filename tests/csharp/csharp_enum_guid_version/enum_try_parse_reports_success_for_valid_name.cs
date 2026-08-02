// vybe-test: csharp/csharp_enum_guid_version/enum_try_parse_reports_success_for_valid_name
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum State { Idle, Running } System.Enum.TryParse<State>("Idle", out var value); __Check((value).ToString(), "Idle");
