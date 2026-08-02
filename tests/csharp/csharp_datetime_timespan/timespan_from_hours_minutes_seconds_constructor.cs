// vybe-test: csharp/csharp_datetime_timespan/timespan_from_hours_minutes_seconds_constructor
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span = new System.TimeSpan(1, 2, 3); __Check((span.TotalSeconds).ToString(), "3723");
