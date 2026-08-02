// vybe-test: csharp/csharp_records_advanced/record_deconstruction_returns_positional_members
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X, int Y); var (x, y) = new Point(8, 9); __Check((x + y).ToString(), "17");
