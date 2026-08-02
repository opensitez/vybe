// vybe-test: csharp/csharp_datetime_offset/datetimeoffset_to_universal_time_yields_utc
// origin: languages/csharp/tests/csharp/test_csharp_datetime_offset.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var dto=new System.DateTimeOffset(2024,1,15,10,0,0,System.TimeSpan.FromHours(2));
var utc=dto.ToUniversalTime();
__Check((utc.Hour).ToString(), "8");
