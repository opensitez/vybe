// vybe-test: csharp/csharp_timespan_arithmetic/timespan_from_seconds_total_seconds
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromSeconds(45); __Check((span.TotalSeconds).ToString(), "45");
