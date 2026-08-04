// vybe-test: csharp/csharp_datetime_timespan/datetime_add_timespan_combines_date_and_duration
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

var date = new System.DateTime(2024, 1, 1, 1, 0, 0); var span = System.TimeSpan.FromMinutes(90); __P(((date + span).Hour).ToString()); __P(((date + span).Minute).ToString());
__Check("2\n30");
