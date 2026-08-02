// vybe-test: csharp/csharp_type_casting/explicit_narrowing_long_to_int_truncates
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

long x = 5L; int y = (int)x; __Check((y).ToString(), "5");
