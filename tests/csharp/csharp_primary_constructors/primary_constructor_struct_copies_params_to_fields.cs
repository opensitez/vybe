// vybe-test: csharp/csharp_primary_constructors/primary_constructor_struct_copies_params_to_fields
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Point(int x, int y) {
    public int X = x;
    public int Y = y;
}
var p = new Point(3, 4);
__Check((p.X).ToString(), "3"); __Check((p.Y).ToString(), "4");
