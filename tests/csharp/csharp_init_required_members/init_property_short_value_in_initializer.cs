// vybe-test: csharp/csharp_init_required_members/init_property_short_value_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class ShortHolder { public short Value { get; init; } }
var s = new ShortHolder { Value = 1000 };
__Check((s.Value).ToString(), "1000");
