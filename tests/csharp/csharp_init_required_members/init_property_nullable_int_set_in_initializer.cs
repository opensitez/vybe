// vybe-test: csharp/csharp_init_required_members/init_property_nullable_int_set_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Maybe { public int? Count { get; init; } }
var m = new Maybe { Count = 5 };
__Check((m.Count).ToString(), "5");
