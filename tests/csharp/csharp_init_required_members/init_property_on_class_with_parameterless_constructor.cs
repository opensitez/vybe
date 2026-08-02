// vybe-test: csharp/csharp_init_required_members/init_property_on_class_with_parameterless_constructor
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Widget { public Widget() { } public int Count { get; init; } = 0; }
var w = new Widget { Count = 5 };
__Check((w.Count).ToString(), "5");
