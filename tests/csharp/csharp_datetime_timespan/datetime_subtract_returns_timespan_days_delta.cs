// vybe-test: csharp/csharp_datetime_timespan/datetime_subtract_returns_timespan_days_delta
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var start = new System.DateTime(2024, 1, 1); var end = new System.DateTime(2024, 1, 4); __Check(((end - start).Days).ToString(), "3");
