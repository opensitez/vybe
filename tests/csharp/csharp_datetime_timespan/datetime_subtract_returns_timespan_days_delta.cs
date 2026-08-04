// vybe-test: csharp/csharp_datetime_timespan/datetime_subtract_returns_timespan_days_delta
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

var start = new System.DateTime(2024, 1, 1); var end = new System.DateTime(2024, 1, 4); __P(((end - start).Days).ToString());
__Check("3");
