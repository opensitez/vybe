// vybe-test: csharp/csharp_records_advanced/record_struct_stores_value_semantics
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Pixel(int X, int Y); var pixel = new Pixel(2, 3); __Check((pixel.X + pixel.Y).ToString(), "5");
