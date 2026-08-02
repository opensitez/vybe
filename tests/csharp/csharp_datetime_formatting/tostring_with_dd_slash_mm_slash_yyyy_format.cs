// vybe-test: csharp/csharp_datetime_formatting/tostring_with_dd_slash_mm_slash_yyyy_format
// origin: languages/csharp/tests/csharp/test_csharp_datetime_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.DateTime(2024,1,5);
__Check((d.ToString("dd/MM/yyyy")).ToString(), "05/01/2024");
