// vybe-test: csharp/csharp_datetime_timespan/datetime_add_days_moves_to_following_date
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var date = new System.DateTime(2024, 1, 30).AddDays(2); __Check((date.Day).ToString(), "1");
