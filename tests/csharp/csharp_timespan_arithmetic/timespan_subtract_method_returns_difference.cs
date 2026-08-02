// vybe-test: csharp/csharp_timespan_arithmetic/timespan_subtract_method_returns_difference
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=System.TimeSpan.FromMinutes(50); var b=System.TimeSpan.FromMinutes(20); __Check((a.Subtract(b).TotalMinutes).ToString(), "30");
