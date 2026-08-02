// vybe-test: csharp/csharp_timespan_arithmetic/timespan_from_days_sets_days_component
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromDays(3); __Check((span.Days).ToString(), "3");
