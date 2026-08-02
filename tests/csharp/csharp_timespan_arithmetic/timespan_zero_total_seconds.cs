// vybe-test: csharp/csharp_timespan_arithmetic/timespan_zero_total_seconds
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.TimeSpan.Zero.TotalSeconds).ToString(), "0");
