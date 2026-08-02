// vybe-test: csharp/csharp_datetime_timespan/datetime_compare_orders_earlier_before_later
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var left = new System.DateTime(2024, 1, 1); var right = new System.DateTime(2024, 1, 2); __Check((System.DateTime.Compare(left, right)).ToString(), "-1");
