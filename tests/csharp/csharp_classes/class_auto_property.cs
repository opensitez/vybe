// vybe-test: csharp/csharp_classes/class_auto_property
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

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
    public string Name { get; set; }
    public int Value { get; set; }
}
var c = new Config();
c.Name = "test";
c.Value = 42;
__P((c.Name).ToString());
__P((c.Value).ToString());
__Check("test\n42");
