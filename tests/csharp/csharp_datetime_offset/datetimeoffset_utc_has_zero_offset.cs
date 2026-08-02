// vybe-test: csharp/csharp_datetime_offset/datetimeoffset_utc_has_zero_offset
// origin: languages/csharp/tests/csharp/test_csharp_datetime_offset.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var dto=System.DateTimeOffset.UtcNow;
__Check((dto.Offset==System.TimeSpan.Zero).ToString(), "True");
