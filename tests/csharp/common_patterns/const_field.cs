// vybe-test: csharp/common_patterns/const_field
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

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
    public const int MaxRetries = 3;
    public const string AppName = "MyApp";
}
__P((Config.MaxRetries).ToString());
__P((Config.AppName).ToString());
__Check("3\nMyApp");
