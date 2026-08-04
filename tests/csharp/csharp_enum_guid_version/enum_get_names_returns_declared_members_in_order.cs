// vybe-test: csharp/csharp_enum_guid_version/enum_get_names_returns_declared_members_in_order
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

enum State { Idle, Running, Done } foreach (var name in System.Enum.GetNames(typeof(State))) __P((name).ToString());
__Check("Idle\nRunning\nDone");
