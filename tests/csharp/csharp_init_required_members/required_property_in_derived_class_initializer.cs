// vybe-test: csharp/csharp_init_required_members/required_property_in_derived_class_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { public int Id; }
class Derived : Base { public required string Label { get; set; } }
var d = new Derived { Label = "child" };
__Check((d.Label).ToString(), "child");
