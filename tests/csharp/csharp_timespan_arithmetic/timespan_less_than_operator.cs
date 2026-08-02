// vybe-test: csharp/csharp_timespan_arithmetic/timespan_less_than_operator
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var left=System.TimeSpan.FromMinutes(2); var right=System.TimeSpan.FromMinutes(5); __Check((left<right).ToString(), "True");
