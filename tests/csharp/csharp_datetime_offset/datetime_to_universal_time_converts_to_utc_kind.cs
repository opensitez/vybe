// vybe-test: csharp/csharp_datetime_offset/datetime_to_universal_time_converts_to_utc_kind
// origin: languages/csharp/tests/csharp/test_csharp_datetime_offset.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var local=new System.DateTime(2024,1,15,12,0,0,System.DateTimeKind.Local);
var utc=local.ToUniversalTime();
__Check((utc.Kind).ToString(), "Utc");
