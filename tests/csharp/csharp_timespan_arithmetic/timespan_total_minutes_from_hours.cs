// vybe-test: csharp/csharp_timespan_arithmetic/timespan_total_minutes_from_hours
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromHours(1.5); __Check((span.TotalMinutes).ToString(), "90");
