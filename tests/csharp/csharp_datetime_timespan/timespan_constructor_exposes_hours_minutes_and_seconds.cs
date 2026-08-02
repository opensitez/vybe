// vybe-test: csharp/csharp_datetime_timespan/timespan_constructor_exposes_hours_minutes_and_seconds
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span = new System.TimeSpan(2, 3, 4); __Check((span.Hours).ToString(), "2"); __Check((span.Minutes).ToString(), "3"); __Check((span.Seconds).ToString(), "4");
