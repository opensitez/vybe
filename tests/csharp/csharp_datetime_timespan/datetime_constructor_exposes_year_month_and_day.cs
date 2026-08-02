// vybe-test: csharp/csharp_datetime_timespan/datetime_constructor_exposes_year_month_and_day
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var date = new System.DateTime(2024, 5, 17); __Check((date.Year).ToString(), "2024"); __Check((date.Month).ToString(), "5"); __Check((date.Day).ToString(), "17");
