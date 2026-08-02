// vybe-test: csharp/csharp_type_conversions/explicit_numeric_conversion_from_double_to_int
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double value = 7.9; int whole = (int)value; __Check((whole).ToString(), "7");
