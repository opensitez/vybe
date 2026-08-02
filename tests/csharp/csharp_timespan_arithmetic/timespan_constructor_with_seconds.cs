// vybe-test: csharp/csharp_timespan_arithmetic/timespan_constructor_with_seconds
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=new System.TimeSpan(0,1,2,3); __Check((span.Hours).ToString(), "1"); __Check((span.Minutes).ToString(), "2"); __Check((span.Seconds).ToString(), "3");
