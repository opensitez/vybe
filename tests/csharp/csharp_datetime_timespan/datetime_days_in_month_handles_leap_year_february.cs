// vybe-test: csharp/csharp_datetime_timespan/datetime_days_in_month_handles_leap_year_february
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.DateTime.DaysInMonth(2024, 2)).ToString(), "29");
