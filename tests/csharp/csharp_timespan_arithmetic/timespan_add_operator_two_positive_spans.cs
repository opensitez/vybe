// vybe-test: csharp/csharp_timespan_arithmetic/timespan_add_operator_two_positive_spans
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=System.TimeSpan.FromDays(1); var b=System.TimeSpan.FromHours(12); __Check(((a+b).TotalHours).ToString(), "36");
