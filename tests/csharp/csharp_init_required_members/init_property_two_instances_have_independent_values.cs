// vybe-test: csharp/csharp_init_required_members/init_property_two_instances_have_independent_values
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Slot { public int Id { get; init; } }
var a = new Slot { Id = 1 };
var b = new Slot { Id = 2 };
__Check((a.Id).ToString(), "1"); __Check((b.Id).ToString(), "2");
