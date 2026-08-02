// vybe-test: csharp/csharp_datetime_timespan/datetime_time_of_day_returns_timespan_component
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var date = new System.DateTime(2024, 1, 1, 6, 45, 0); __Check((date.TimeOfDay.TotalMinutes).ToString(), "405");
