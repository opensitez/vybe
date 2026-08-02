// vybe-test: csharp/csharp_timespan_arithmetic/timespan_from_hours_negative_total
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromHours(-1); __Check((span.TotalHours).ToString(), "-1");
