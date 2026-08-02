// vybe-test: csharp/csharp_type_conversions/implicit_numeric_conversion_from_int_to_double
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int count = 7; double total = count; __Check((total).ToString(), "7");
