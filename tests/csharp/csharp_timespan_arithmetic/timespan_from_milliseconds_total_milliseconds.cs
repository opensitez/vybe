// vybe-test: csharp/csharp_timespan_arithmetic/timespan_from_milliseconds_total_milliseconds
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromMilliseconds(1500); __Check((span.TotalMilliseconds).ToString(), "1500");
