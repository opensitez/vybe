// vybe-test: csharp/csharp_timespan_arithmetic/timespan_constructor_days_hours_minutes
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=new System.TimeSpan(1,2,30); __Check((span.Days).ToString(), "1"); __Check((span.Hours).ToString(), "2"); __Check((span.Minutes).ToString(), "30");
