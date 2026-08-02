// vybe-test: csharp/csharp_timespan_arithmetic/timespan_from_days_total_days
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromDays(2.5); __Check((span.TotalDays).ToString(), "2.5");
