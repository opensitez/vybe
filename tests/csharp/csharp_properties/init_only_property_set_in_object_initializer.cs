// vybe-test: csharp/csharp_properties/init_only_property_set_in_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Point { public int X { get; init; } public int Y { get; init; } }
var p = new Point { X=1, Y=2 };
__Check((p.X).ToString(), "1"); __Check((p.Y).ToString(), "2");
