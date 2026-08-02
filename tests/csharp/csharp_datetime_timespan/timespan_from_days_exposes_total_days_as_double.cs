// vybe-test: csharp/csharp_datetime_timespan/timespan_from_days_exposes_total_days_as_double
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span = System.TimeSpan.FromDays(2); __Check((span.TotalDays).ToString(), "2");
