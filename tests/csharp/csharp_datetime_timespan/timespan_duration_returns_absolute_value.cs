// vybe-test: csharp/csharp_datetime_timespan/timespan_duration_returns_absolute_value
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span = System.TimeSpan.FromSeconds(-12).Duration(); __Check((span.TotalSeconds).ToString(), "12");
