// vybe-test: csharp/csharp_timespan_arithmetic/timespan_compare_to_equal_is_zero
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var left=System.TimeSpan.FromSeconds(5); var right=System.TimeSpan.FromSeconds(5); __Check((left.CompareTo(right)).ToString(), "0");
