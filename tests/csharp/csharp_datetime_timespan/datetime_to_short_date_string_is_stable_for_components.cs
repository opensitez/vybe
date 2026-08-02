// vybe-test: csharp/csharp_datetime_timespan/datetime_to_short_date_string_is_stable_for_components
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var date = new System.DateTime(2024, 12, 25); var text = date.ToShortDateString(); __Check((text.Contains("2024")).ToString(), "True");
