// vybe-test: csharp/csharp_timespan_arithmetic/timespan_static_compare_matches_instance
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var left=System.TimeSpan.FromHours(1); var right=System.TimeSpan.FromHours(2); __Check((System.TimeSpan.Compare(left,right)).ToString(), "-1");
