// vybe-test: csharp/csharp_timespan_arithmetic/timespan_add_method_combines_durations
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=System.TimeSpan.FromHours(1); var b=System.TimeSpan.FromMinutes(30); __Check((a.Add(b).TotalMinutes).ToString(), "90");
