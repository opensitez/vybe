// vybe-test: csharp/csharp_datetime_formatting/tostring_with_hh_mm_ss_time_format
// origin: languages/csharp/tests/csharp/test_csharp_datetime_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.DateTime(2024,1,1,13,5,9);
__Check((d.ToString("HH:mm:ss")).ToString(), "13:05:09");
