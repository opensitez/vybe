// vybe-test: csharp/csharp_numeric_checked_bitwise/compound_shift_assignment_updates_variable_in_place
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 5; value <<= 1; __Check((value).ToString(), "10");
