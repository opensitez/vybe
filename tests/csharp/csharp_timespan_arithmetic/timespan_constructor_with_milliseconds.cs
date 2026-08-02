// vybe-test: csharp/csharp_timespan_arithmetic/timespan_constructor_with_milliseconds
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=new System.TimeSpan(0,0,0,0,250); __Check((span.Milliseconds).ToString(), "250");
