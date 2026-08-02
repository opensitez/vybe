// vybe-test: csharp/csharp_timespan_arithmetic/timespan_compare_negative_to_positive
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var left=System.TimeSpan.FromMinutes(-5); var right=System.TimeSpan.FromMinutes(5); __Check((left.CompareTo(right)).ToString(), "-1");
