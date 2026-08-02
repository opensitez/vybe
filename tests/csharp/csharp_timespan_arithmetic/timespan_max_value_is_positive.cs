// vybe-test: csharp/csharp_timespan_arithmetic/timespan_max_value_is_positive
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.TimeSpan.MaxValue.TotalDays>0).ToString(), "True");
