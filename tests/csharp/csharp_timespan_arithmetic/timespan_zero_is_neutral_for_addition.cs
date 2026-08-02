// vybe-test: csharp/csharp_timespan_arithmetic/timespan_zero_is_neutral_for_addition
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromMinutes(10); __Check((span.Add(System.TimeSpan.Zero).TotalMinutes).ToString(), "10");
