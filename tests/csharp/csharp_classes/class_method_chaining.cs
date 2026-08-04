// vybe-test: csharp/csharp_classes/class_method_chaining
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

class Builder {
    private string result = "";
    public Builder Add(string s) {
        result += s;
        return this;
    }
    public string Build() { return result; }
}
var r = new Builder().Add("a").Add("b").Add("c").Build();
__P((r).ToString());
__Check("abc");
