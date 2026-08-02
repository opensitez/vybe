// vybe-test: csharp/csharp_init_required_members/init_property_on_struct_with_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Pair { public int A { get; init; } public int B { get; init; } }
var p = new Pair { A = 4, B = 6 };
__Check((p.A + p.B).ToString(), "10");
