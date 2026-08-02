// vybe-test: csharp/csharp_datetime_formatting/tostring_with_yyyy_mm_dd_format
// origin: languages/csharp/tests/csharp/test_csharp_datetime_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.DateTime(2024,6,15);
__Check((d.ToString("yyyy-MM-dd")).ToString(), "2024-06-15");
