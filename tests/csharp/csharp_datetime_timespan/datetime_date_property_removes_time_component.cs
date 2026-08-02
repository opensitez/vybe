// vybe-test: csharp/csharp_datetime_timespan/datetime_date_property_removes_time_component
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var date = new System.DateTime(2024, 7, 8, 13, 14, 15).Date; __Check((date.Hour).ToString(), "0");
