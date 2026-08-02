// vybe-test: csharp/csharp_init_required_members/init_property_double_set_in_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Measure { public double Value { get; init; } }
var m = new Measure { Value = 3.5 };
__Check((m.Value).ToString(), "3.5");
