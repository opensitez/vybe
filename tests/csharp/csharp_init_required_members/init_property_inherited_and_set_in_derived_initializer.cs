// vybe-test: csharp/csharp_init_required_members/init_property_inherited_and_set_in_derived_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { public string Tag { get; init; } = "base"; }
class Derived : Base { }
var d = new Derived { Tag = "child" };
__Check((d.Tag).ToString(), "child");
