// vybe-test: csharp/csharp_datetime_timespan/timespan_addition_combines_hours_and_minutes
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var left = System.TimeSpan.FromHours(1); var right = System.TimeSpan.FromMinutes(30); __P(((left + right).TotalMinutes).ToString());
__Check("90");
