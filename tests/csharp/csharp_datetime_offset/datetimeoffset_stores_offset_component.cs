// vybe-test: csharp/csharp_datetime_offset/datetimeoffset_stores_offset_component
// origin: languages/csharp/tests/csharp/test_csharp_datetime_offset.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var dto=new System.DateTimeOffset(2024,1,15,10,0,0,System.TimeSpan.FromHours(5));
__Check((dto.Offset.Hours).ToString(), "5");
