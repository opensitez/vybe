// vybe-test: csharp/csharp_timespan_arithmetic/timespan_add_negative_span_via_negate
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var baseSpan=System.TimeSpan.FromHours(2); var delta=System.TimeSpan.FromMinutes(30); __Check((baseSpan.Add(-delta).TotalMinutes).ToString(), "90");
