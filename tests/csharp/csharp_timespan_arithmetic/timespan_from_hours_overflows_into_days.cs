// vybe-test: csharp/csharp_timespan_arithmetic/timespan_from_hours_overflows_into_days
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromHours(25); __Check((span.Days).ToString(), "1"); __Check((span.Hours).ToString(), "1");
