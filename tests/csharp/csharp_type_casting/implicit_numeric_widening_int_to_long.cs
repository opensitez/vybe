// vybe-test: csharp/csharp_type_casting/implicit_numeric_widening_int_to_long
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 100; long y = x; __Check((y).ToString(), "100");
