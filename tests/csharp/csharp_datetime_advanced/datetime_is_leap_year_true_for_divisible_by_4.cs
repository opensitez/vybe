// vybe-test: csharp/csharp_datetime_advanced/datetime_is_leap_year_true_for_divisible_by_4
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

__P((System.DateTime.IsLeapYear(2024)).ToString());
__P((System.DateTime.IsLeapYear(2023)).ToString());
__Check("True\nFalse");
