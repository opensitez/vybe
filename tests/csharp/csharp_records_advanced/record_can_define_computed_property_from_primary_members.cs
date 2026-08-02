// vybe-test: csharp/csharp_records_advanced/record_can_define_computed_property_from_primary_members
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Rectangle(int Width, int Height) { public int Area => Width * Height; } __Check((new Rectangle(3, 7).Area).ToString(), "21");
