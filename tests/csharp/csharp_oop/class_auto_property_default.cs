// vybe-test: csharp/csharp_oop/class_auto_property_default
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

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
    public string Name { get; set; } = "default";
    public int Count { get; set; } = 0;
}
var c = new Config();
__P((c.Name).ToString());
c.Name = "custom";
__P((c.Name).ToString());
__Check("default\ncustom");
