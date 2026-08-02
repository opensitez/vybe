// vybe-test: csharp/oop_advanced/this_reference_return
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((b.Add("A").Add("B").Add("C").Build()).ToString(), "[A, B, C]");
