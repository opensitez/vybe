// vybe-test: csharp/csharp_timespan_arithmetic/timespan_to_string_positive_hms
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.TimeSpan.FromHours(1).Add(System.TimeSpan.FromMinutes(2)).Add(System.TimeSpan.FromSeconds(3)).ToString()).ToString(), "01:02:03");
