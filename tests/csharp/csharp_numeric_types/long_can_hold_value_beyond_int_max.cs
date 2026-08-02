// vybe-test: csharp/csharp_numeric_types/long_can_hold_value_beyond_int_max
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

long x = (long)int.MaxValue + 1; __Check((x).ToString(), "2147483648");
