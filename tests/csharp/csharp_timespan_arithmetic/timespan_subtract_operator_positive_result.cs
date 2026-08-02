// vybe-test: csharp/csharp_timespan_arithmetic/timespan_subtract_operator_positive_result
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=System.TimeSpan.FromHours(3); var b=System.TimeSpan.FromHours(1); __Check(((a-b).TotalHours).ToString(), "2");
