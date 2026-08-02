// vybe-test: csharp/csharp_timespan_arithmetic/timespan_add_multiple_increments
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.Zero; span=span.Add(System.TimeSpan.FromMinutes(10)); span=span.Add(System.TimeSpan.FromMinutes(5)); __Check((span.TotalMinutes).ToString(), "15");
