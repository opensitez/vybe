// vybe-test: csharp/csharp_datetime_timespan/datetime_kind_utc_survives_property_read
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var instant = new System.DateTime(2024, 6, 1, 0, 0, 0, System.DateTimeKind.Utc); __Check((instant.Kind).ToString(), "Utc");
