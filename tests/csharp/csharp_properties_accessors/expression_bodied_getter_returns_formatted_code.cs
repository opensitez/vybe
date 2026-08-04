// vybe-test: csharp/csharp_properties_accessors/expression_bodied_getter_returns_formatted_code
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

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

class Package {
    public string Prefix { get; set; }
    public int Number { get; set; }
    public string Code => Prefix + "-" + Number;
}
var package = new Package { Prefix = "PKG", Number = 42 };
__P((package.Code).ToString());
__Check("PKG-42");
