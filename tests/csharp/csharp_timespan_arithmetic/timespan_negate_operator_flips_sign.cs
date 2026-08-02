// vybe-test: csharp/csharp_timespan_arithmetic/timespan_negate_operator_flips_sign
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var span=System.TimeSpan.FromMinutes(15); __Check(((-span).TotalMinutes).ToString(), "-15");
