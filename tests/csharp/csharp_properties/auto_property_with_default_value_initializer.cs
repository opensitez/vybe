// vybe-test: csharp/csharp_properties/auto_property_with_default_value_initializer
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

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

class Config { public int Timeout { get; set; } = 30; }
__P((new Config().Timeout).ToString());
__Check("30");
