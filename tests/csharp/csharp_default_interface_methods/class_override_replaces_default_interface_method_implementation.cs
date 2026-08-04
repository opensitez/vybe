// vybe-test: csharp/csharp_default_interface_methods/class_override_replaces_default_interface_method_implementation
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods.rs

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

interface IFormat {
    string Format(int n);
    string Label(int n) { return "d:" + Format(n); }
}
class Custom : IFormat {
    public string Format(int n) { return n.ToString(); }
    public string Label(int n) { return "x:" + Format(n); }
}
IFormat fmt = new Custom();
__P((fmt.Label(3)).ToString());
__Check("x:3");
