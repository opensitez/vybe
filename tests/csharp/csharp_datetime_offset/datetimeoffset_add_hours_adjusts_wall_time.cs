// vybe-test: csharp/csharp_datetime_offset/datetimeoffset_add_hours_adjusts_wall_time
// origin: languages/csharp/tests/csharp/test_csharp_datetime_offset.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var dto=new System.DateTimeOffset(2024,1,1,20,0,0,System.TimeSpan.Zero);
var next=dto.AddHours(5);
__Check((next.Day).ToString(), "2"); __Check((next.Hour).ToString(), "1");
