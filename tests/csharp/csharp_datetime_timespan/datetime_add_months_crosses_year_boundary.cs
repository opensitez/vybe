// vybe-test: csharp/csharp_datetime_timespan/datetime_add_months_crosses_year_boundary
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var date = new System.DateTime(2023, 11, 15).AddMonths(3); __Check((date.Year).ToString(), "2024"); __Check((date.Month).ToString(), "2");
