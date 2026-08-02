// vybe-test: csharp/csharp_records_advanced/record_hash_code_matches_for_equal_values
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X, int Y); var left = new Point(5, 7); var right = new Point(5, 7); __Check((left.GetHashCode() == right.GetHashCode()).ToString(), "True");
