// vybe-test: csharp/csharp_timespan_arithmetic/timespan_compare_to_shorter_is_negative
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var left=System.TimeSpan.FromMinutes(1); var right=System.TimeSpan.FromMinutes(2); __Check((left.CompareTo(right)).ToString(), "-1");
