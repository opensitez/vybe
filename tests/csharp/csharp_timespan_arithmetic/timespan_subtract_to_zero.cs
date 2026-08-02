// vybe-test: csharp/csharp_timespan_arithmetic/timespan_subtract_to_zero
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromMinutes(10); __Check((span.Subtract(span).TotalMinutes).ToString(), "0");
