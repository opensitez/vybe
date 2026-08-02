// vybe-test: csharp/csharp_timespan_arithmetic/timespan_from_ticks_reads_ticks_property
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromTicks(10000000); __Check((span.Ticks).ToString(), "10000000");
