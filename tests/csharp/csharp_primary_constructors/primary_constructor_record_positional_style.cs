// vybe-test: csharp/csharp_primary_constructors/primary_constructor_record_positional_style
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X, int Y);
var p = new Point(1, 2);
__Check((p.X).ToString(), "1"); __Check((p.Y).ToString(), "2");
