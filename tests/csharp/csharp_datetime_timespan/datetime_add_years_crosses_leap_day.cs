// vybe-test: csharp/csharp_datetime_timespan/datetime_add_years_crosses_leap_day
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var date = new System.DateTime(2024, 2, 29).AddYears(1); __Check((date.Month).ToString(), "2"); __Check((date.Day).ToString(), "28");
