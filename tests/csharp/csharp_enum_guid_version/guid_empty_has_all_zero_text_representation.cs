// vybe-test: csharp/csharp_enum_guid_version/guid_empty_has_all_zero_text_representation
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

__P((System.Guid.Empty.ToString()).ToString());
__Check("00000000-0000-0000-0000-000000000000");
