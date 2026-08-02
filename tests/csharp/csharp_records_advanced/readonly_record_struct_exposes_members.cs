// vybe-test: csharp/csharp_records_advanced/readonly_record_struct_exposes_members
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly record struct Size(int Width, int Height); var size = new Size(4, 6); __Check((size.Width * size.Height).ToString(), "24");
