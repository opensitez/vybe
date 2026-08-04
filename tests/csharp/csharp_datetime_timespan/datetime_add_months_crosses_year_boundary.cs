// vybe-test: csharp/csharp_datetime_timespan/datetime_add_months_crosses_year_boundary
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

var date = new System.DateTime(2023, 11, 15).AddMonths(3); __P((date.Year).ToString()); __P((date.Month).ToString());
__Check("2024\n2");
