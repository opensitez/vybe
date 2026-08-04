// vybe-test: csharp/csharp_enum_guid_version/guid_try_parse_reports_true_for_valid_text
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

var ok = System.Guid.TryParse("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee", out var value); __P((ok).ToString()); __P((value.ToString()).ToString());
__Check("True\naaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee");
