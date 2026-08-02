// vybe-test: csharp/csharp_datetime_timespan/timespan_subtraction_returns_remaining_minutes
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var left = System.TimeSpan.FromMinutes(45); var right = System.TimeSpan.FromMinutes(5); __Check(((left - right).TotalMinutes).ToString(), "40");
