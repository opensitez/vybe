// vybe-test: csharp/csharp_init_required_members/required_property_with_init_accessor
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Node { public required int Id { get; init; } }
var n = new Node { Id = 42 };
__Check((n.Id).ToString(), "42");
