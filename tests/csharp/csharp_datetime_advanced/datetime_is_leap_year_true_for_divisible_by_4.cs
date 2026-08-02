// vybe-test: csharp/csharp_datetime_advanced/datetime_is_leap_year_true_for_divisible_by_4
// origin: languages/csharp/tests/csharp/test_csharp_datetime_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.DateTime.IsLeapYear(2024)).ToString(), "True");
__Check((System.DateTime.IsLeapYear(2023)).ToString(), "False");
