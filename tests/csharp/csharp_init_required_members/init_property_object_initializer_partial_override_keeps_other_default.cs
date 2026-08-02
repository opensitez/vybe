// vybe-test: csharp/csharp_init_required_members/init_property_object_initializer_partial_override_keeps_other_default
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pair { public int A { get; init; } = 1; public int B { get; init; } = 2; }
var p = new Pair { B = 9 };
__Check((p.A).ToString(), "1"); __Check((p.B).ToString(), "9");
