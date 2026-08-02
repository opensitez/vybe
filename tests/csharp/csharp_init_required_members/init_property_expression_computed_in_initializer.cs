// vybe-test: csharp/csharp_init_required_members/init_property_expression_computed_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { public int Size { get; init; } }
var b = new Box { Size = 10 + 5 };
__Check((b.Size).ToString(), "15");
