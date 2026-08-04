// vybe-test: csharp/csharp_enum_guid_version/version_parse_exposes_major_minor_and_build
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

var version = System.Version.Parse("2.4.6"); __P((version.Major).ToString()); __P((version.Minor).ToString()); __P((version.Build).ToString());
__Check("2\n4\n6");
