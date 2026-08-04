// vybe-test: csharp/csharp_guid_parse_format/guid_parse_accepts_standard_hyphenated_representation
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
__P((id.ToString().StartsWith("11111111")).ToString());
__Check("True");
