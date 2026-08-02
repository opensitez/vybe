// vybe-test: csharp/csharp_numeric_types/uint_holds_value_beyond_int_max
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

uint u = 3000000000u; __Check((u > int.MaxValue).ToString(), "True");
