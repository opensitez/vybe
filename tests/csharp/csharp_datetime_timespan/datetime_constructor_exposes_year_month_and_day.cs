// vybe-test: csharp/csharp_datetime_timespan/datetime_constructor_exposes_year_month_and_day
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

var date = new System.DateTime(2024, 5, 17); __P((date.Year).ToString()); __P((date.Month).ToString()); __P((date.Day).ToString());
__Check("2024\n5\n17");
