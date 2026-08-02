// vybe-test: csharp/csharp_init_required_members/required_and_init_on_same_property_via_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Entity { public required int Id { get; init; } }
var e = new Entity { Id = 7 };
__Check((e.Id).ToString(), "7");
