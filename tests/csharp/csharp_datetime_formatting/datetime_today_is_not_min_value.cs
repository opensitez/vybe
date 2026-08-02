// vybe-test: csharp/csharp_datetime_formatting/datetime_today_is_not_min_value
// origin: languages/csharp/tests/csharp/test_csharp_datetime_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.DateTime.Today != System.DateTime.MinValue).ToString(), "True");
