// vybe-test: csharp/csharp_datetime_timespan/datetime_is_leap_year_recognizes_2024
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.DateTime.IsLeapYear(2024)).ToString(), "True");
