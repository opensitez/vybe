// vybe-test: csharp/oop_advanced/this_reference_return
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

class Builder {
    string parts = "";
    public Builder Add(string part) {
        if (parts.Length > 0) parts += ", ";
        parts += part;
        return this;
    }
    public string Build() { return "[" + parts + "]"; }
}
var b = new Builder();
__P((b.Add("A").Add("B").Add("C").Build()).ToString());
__Check("[A, B, C]");
