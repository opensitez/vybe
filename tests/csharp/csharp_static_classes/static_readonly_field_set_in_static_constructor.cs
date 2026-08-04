// vybe-test: csharp/csharp_static_classes/static_readonly_field_set_in_static_constructor
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

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

class Config {
    public static readonly string Version;
    static Config() { Version = "1.0"; }
}
__P((Config.Version).ToString());
__Check("1.0");
