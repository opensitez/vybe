// vybe-test: csharp/csharp_datetime_timespan/timespan_addition_combines_hours_and_minutes
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var left = System.TimeSpan.FromHours(1); var right = System.TimeSpan.FromMinutes(30); __Check(((left + right).TotalMinutes).ToString(), "90");
