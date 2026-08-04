// vybe-test: csharp/csharp_nested_partial_types/partial_class_combines_field_and_method_declarations
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

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

partial class Config {
    string env = "prod";
}
partial class Config {
    public string Read() { return env; }
}
__P((new Config().Read()).ToString());
__Check("prod");
