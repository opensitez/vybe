// vybe-test: csharp/csharp_timespan_arithmetic/timespan_subtract_self_compare_zero
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromDays(1); __Check((span.Subtract(span).CompareTo(System.TimeSpan.Zero)).ToString(), "0");
