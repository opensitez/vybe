// vybe-test: csharp/csharp_primary_constructors/primary_constructor_record_struct_value_type
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Coord(int X, int Y);
var c = new Coord(5, 6);
__Check((c.X + c.Y).ToString(), "11");
