// vybe-test: csharp/csharp_records_advanced/positional_record_to_string_includes_member_values
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X, int Y); __Check((new Point(3, 4).ToString().Contains("X = 3")).ToString(), "True");
