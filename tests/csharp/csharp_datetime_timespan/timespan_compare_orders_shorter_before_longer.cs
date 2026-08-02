// vybe-test: csharp/csharp_datetime_timespan/timespan_compare_orders_shorter_before_longer
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var left = System.TimeSpan.FromSeconds(3); var right = System.TimeSpan.FromSeconds(8); __Check((System.TimeSpan.Compare(left, right)).ToString(), "-1");
