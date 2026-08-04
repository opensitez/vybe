// vybe-test: csharp/csharp_guid_parse_format/guid_to_string_with_format_specifier_renders_hyphenated_value
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_format.rs

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

var id = System.Guid.Parse("11111111-2222-3333-4444-555555555555");
__P((id.ToString("D").StartsWith("11111111")).ToString());
__Check("True");
