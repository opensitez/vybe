// vybe-test: csharp/csharp_datetime_timespan/timespan_constructor_exposes_hours_minutes_and_seconds
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

var span = new System.TimeSpan(2, 3, 4); __P((span.Hours).ToString()); __P((span.Minutes).ToString()); __P((span.Seconds).ToString());
__Check("2\n3\n4");
