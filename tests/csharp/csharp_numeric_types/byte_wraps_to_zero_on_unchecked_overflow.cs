// vybe-test: csharp/csharp_numeric_types/byte_wraps_to_zero_on_unchecked_overflow
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

unchecked { byte b = 255; b++; __Check((b).ToString(), "0"); }
