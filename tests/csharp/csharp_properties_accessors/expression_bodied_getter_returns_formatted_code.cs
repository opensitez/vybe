// vybe-test: csharp/csharp_properties_accessors/expression_bodied_getter_returns_formatted_code
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

using static __Harness;

var package = new Package { Prefix = "PKG", Number = 42 }
;
__P((package.Code).ToString());
__Check("PKG-42");

class Package {
    public string Prefix { get; set; }
    public int Number { get; set; }
    public string Code => Prefix + "-" + Number;
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
