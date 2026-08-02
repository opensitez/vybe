// vybe-test: csharp/csharp_datetime_timespan/datetime_day_of_week_reports_expected_enum_name
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var date = new System.DateTime(2024, 6, 3); __Check((date.DayOfWeek).ToString(), "Monday");
