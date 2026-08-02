// vybe-test: csharp/csharp_init_required_members/init_property_multiple_on_same_type
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Point { public int X { get; init; } public int Y { get; init; } }
var p = new Point { X = 2, Y = 7 };
__Check((p.X).ToString(), "2"); __Check((p.Y).ToString(), "7");
