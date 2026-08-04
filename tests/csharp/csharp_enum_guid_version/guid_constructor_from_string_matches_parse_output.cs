// vybe-test: csharp/csharp_enum_guid_version/guid_constructor_from_string_matches_parse_output
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

var text = "11111111-2222-3333-4444-555555555555"; __P((new System.Guid(text).ToString()).ToString());
__Check("11111111-2222-3333-4444-555555555555");
