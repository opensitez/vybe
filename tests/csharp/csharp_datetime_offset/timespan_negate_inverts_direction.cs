// vybe-test: csharp/csharp_datetime_offset/timespan_negate_inverts_direction
// origin: languages/csharp/tests/csharp/test_csharp_datetime_offset.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var ts=System.TimeSpan.FromHours(3);
__Check(((-ts).Hours).ToString(), "-3");
