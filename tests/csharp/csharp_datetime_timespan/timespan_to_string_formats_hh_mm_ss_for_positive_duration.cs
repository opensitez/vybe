// vybe-test: csharp/csharp_datetime_timespan/timespan_to_string_formats_hh_mm_ss_for_positive_duration
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.TimeSpan.FromSeconds(5).ToString()).ToString(), "00:00:05");
