// vybe-test: csharp/csharp_datetime_advanced/datetime_days_in_month_february_leap
// origin: languages/csharp/tests/csharp/test_csharp_datetime_advanced.rs

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

__P((System.DateTime.DaysInMonth(2024,2)).ToString());
__P((System.DateTime.DaysInMonth(2023,2)).ToString());
__Check("29\n28");
