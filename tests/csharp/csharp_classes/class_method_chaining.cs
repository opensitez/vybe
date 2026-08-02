// vybe-test: csharp/csharp_classes/class_method_chaining
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((r).ToString(), "abc");
