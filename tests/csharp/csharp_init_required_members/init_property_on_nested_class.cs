// vybe-test: csharp/csharp_init_required_members/init_property_on_nested_class
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer { public class Inner { public string Name { get; init; } } }
var i = new Outer.Inner { Name = "core" };
__Check((i.Name).ToString(), "core");
