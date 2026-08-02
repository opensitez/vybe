// vybe-test: csharp/csharp_timespan_arithmetic/timespan_total_seconds_from_milliseconds
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromMilliseconds(2500); __Check((span.TotalSeconds).ToString(), "2.5");
