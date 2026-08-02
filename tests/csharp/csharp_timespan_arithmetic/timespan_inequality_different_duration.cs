// vybe-test: csharp/csharp_timespan_arithmetic/timespan_inequality_different_duration
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=System.TimeSpan.FromMinutes(60); var b=System.TimeSpan.FromHours(2); __Check((a!=b).ToString(), "True");
