// vybe-test: csharp/csharp_datetime_formatting/tostring_d_short_date_pattern_contains_year_digits
// origin: languages/csharp/tests/csharp/test_csharp_datetime_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.DateTime(2025,12,31);
__Check((d.ToString("yyyy-MM-dd").StartsWith("2025")).ToString(), "True");
