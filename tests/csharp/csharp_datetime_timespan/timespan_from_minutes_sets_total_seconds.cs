// vybe-test: csharp/csharp_datetime_timespan/timespan_from_minutes_sets_total_seconds
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span = System.TimeSpan.FromMinutes(2.5); __Check((span.TotalSeconds).ToString(), "150");
