// vybe-test: csharp/csharp_timespan_arithmetic/timespan_add_commutative_via_total_minutes
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=System.TimeSpan.FromMinutes(10); var b=System.TimeSpan.FromMinutes(20); __Check(((a+b).TotalMinutes).ToString(), "30"); __Check(((b+a).TotalMinutes).ToString(), "30");
