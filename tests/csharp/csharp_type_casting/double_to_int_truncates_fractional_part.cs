// vybe-test: csharp/csharp_type_casting/double_to_int_truncates_fractional_part
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double d = 9.9; int n = (int)d; __Check((n).ToString(), "9");
