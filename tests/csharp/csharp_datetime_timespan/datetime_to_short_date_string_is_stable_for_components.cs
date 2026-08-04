// vybe-test: csharp/csharp_datetime_timespan/datetime_to_short_date_string_is_stable_for_components
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

var date = new System.DateTime(2024, 12, 25); var text = date.ToShortDateString(); __P((text.Contains("2024")).ToString());
__Check("True");
