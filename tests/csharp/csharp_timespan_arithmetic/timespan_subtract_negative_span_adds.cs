// vybe-test: csharp/csharp_timespan_arithmetic/timespan_subtract_negative_span_adds
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var baseSpan=System.TimeSpan.FromHours(1); var delta=System.TimeSpan.FromMinutes(-30); __Check((baseSpan.Subtract(delta).TotalMinutes).ToString(), "90");
