// vybe-test: csharp/csharp_timespan_arithmetic/timespan_from_hours_sets_hours_component
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromHours(5); __Check((span.Hours).ToString(), "5");
