// vybe-test: csharp/csharp_datetime_advanced/datetime_days_in_month_february_leap
// origin: languages/csharp/tests/csharp/test_csharp_datetime_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.DateTime.DaysInMonth(2024,2)).ToString(), "29");
__Check((System.DateTime.DaysInMonth(2023,2)).ToString(), "28");
