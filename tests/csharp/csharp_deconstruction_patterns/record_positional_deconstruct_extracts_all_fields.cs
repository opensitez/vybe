// vybe-test: csharp/csharp_deconstruction_patterns/record_positional_deconstruct_extracts_all_fields
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X, int Y, int Z);
var p = new Point(1,2,3);
var (x,y,z) = p;
__Check((x+y+z).ToString(), "6");
