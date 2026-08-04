// vybe-test: csharp/csharp_guid_parse_format/guid_empty_has_all_zero_bytes
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

var id = System.Guid.Empty;
__P((id == new System.Guid("00000000-0000-0000-0000-000000000000")).ToString());
__Check("True");
