// vybe-test: csharp/csharp_init_required_members/init_property_float_value_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Sample { public float Rate { get; init; } }
var s = new Sample { Rate = 2.5f };
__Check((s.Rate).ToString(), "2.5");
