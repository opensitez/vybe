// vybe-test: csharp/oop_advanced/sealed_class_basic
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

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

sealed class Config {
    public string Name { get; set; }
    public Config(string n) { Name = n; }
}
var c = new Config("prod");
__P((c.Name).ToString());
__Check("prod");
