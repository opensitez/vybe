// vybe-test: csharp/csharp_datetime_timespan/datetime_add_timespan_combines_date_and_duration
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var date = new System.DateTime(2024, 1, 1, 1, 0, 0); var span = System.TimeSpan.FromMinutes(90); __Check(((date + span).Hour).ToString(), "2"); __Check(((date + span).Minute).ToString(), "30");
