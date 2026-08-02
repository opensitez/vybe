// vybe-test: csharp/csharp_timespan_arithmetic/timespan_duration_of_positive_span_unchanged
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromHours(2); __Check((span.Duration().TotalHours).ToString(), "2");
