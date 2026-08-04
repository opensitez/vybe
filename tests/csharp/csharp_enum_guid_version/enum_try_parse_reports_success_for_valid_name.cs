// vybe-test: csharp/csharp_enum_guid_version/enum_try_parse_reports_success_for_valid_name
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

enum State { Idle, Running } System.Enum.TryParse<State>("Idle", out var value); __P((value).ToString());
__Check("Idle");
