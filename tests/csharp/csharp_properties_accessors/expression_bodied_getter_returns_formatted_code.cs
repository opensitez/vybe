// vybe-test: csharp/csharp_properties_accessors/expression_bodied_getter_returns_formatted_code
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Package {
    public string Prefix { get; set; }
    public int Number { get; set; }
    public string Code => Prefix + "-" + Number;
}
var package = new Package { Prefix = "PKG", Number = 42 };
__Check((package.Code).ToString(), "PKG-42");
