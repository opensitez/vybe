// vybe-test: csharp/csharp_datetime_advanced/timespan_total_minutes_converts_hours_and_minutes
// origin: languages/csharp/tests/csharp/test_csharp_datetime_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var ts=new System.TimeSpan(2,30,0);
__Check((ts.TotalMinutes).ToString(), "150");
