// vybe-test: csharp/csharp_tuples_ranges/double_parse
// origin: languages/csharp/tests/csharp/test_csharp_tuples_ranges.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double d = double.Parse("3.14");
__Check((d).ToString(), "3.14");
