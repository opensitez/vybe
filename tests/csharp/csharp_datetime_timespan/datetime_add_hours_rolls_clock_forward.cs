// vybe-test: csharp/csharp_datetime_timespan/datetime_add_hours_rolls_clock_forward
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var date = new System.DateTime(2024, 1, 1, 10, 30, 0).AddHours(5); __Check((date.Hour).ToString(), "15");
