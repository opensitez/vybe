// vybe-test: csharp/csharp_records_advanced/record_struct_equality_compares_member_values
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Pixel(int X, int Y); __Check((new Pixel(1, 1) == new Pixel(1, 1)).ToString(), "True");
