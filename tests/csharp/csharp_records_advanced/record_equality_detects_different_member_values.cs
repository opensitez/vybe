// vybe-test: csharp/csharp_records_advanced/record_equality_detects_different_member_values
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X, int Y); __Check((new Point(1, 2) == new Point(2, 1)).ToString(), "False");
