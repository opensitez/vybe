// vybe-test: csharp/csharp_datetime_timespan/datetime_ticks_counts_interval_since_epoch_origin
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var epoch = new System.DateTime(1970, 1, 1, 0, 0, 0, System.DateTimeKind.Utc); __Check((epoch.Ticks > 0).ToString(), "True");
